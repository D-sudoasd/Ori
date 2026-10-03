"""Tk interface for importing general tabular data into Origin or Excel."""

from __future__ import annotations

import copy
import json
import multiprocessing
import os
import queue
import re
import subprocess
import sys
import threading
import time
from collections import deque
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
from typing import Any

from .batch_config import (
    config_path_for,
    find_config_for_record,
    load_batch_config,
    save_batch_config,
)
from .readers import SUPPORTED_EXTENSIONS, _natural_key, discover_files
from .source_summary import SummaryStore, options_fingerprint
from .worker import batch_worker, import_worker


SUPPORTED_SUFFIXES = SUPPORTED_EXTENSIONS
PLAN_SCHEMA_VERSION = 1
_FORMAT_SUFFIX = {"opju": ".opju", "xlsx": ".xlsx"}
_DELIMITER_OPTIONS = ("自动识别", "逗号 ,", "制表符 \\t", "分号 ;", "空白分隔")
# Main-thread budgets. Folder discovery and per-source inspect run off this thread.
_UI_FRAME_SECONDS = 0.012
_UI_ROW_BUDGET = 20
_UI_MESSAGE_BUDGET = 20
_IMPORT_MESSAGE_BUDGET = 8
_IMPORT_FRAME_SECONDS = 0.012
_IMPORT_EXIT_POLLS = 2


def _normalized_path(path: str | Path) -> str:
    return os.path.normcase(str(Path(path).expanduser().resolve(strict=False)))


def _source_key(source: dict[str, Any], base_dir: Path | None = None) -> tuple[str, str | None]:
    options = source.get("options") or {}
    sheet = options.get("sheet", source.get("sheet"))
    return _normalized_path(_resolve_plan_path(source["path"], base_dir)), str(sheet) if sheet is not None else None


def _plan_table_key(table: dict[str, Any], base_dir: Path | None = None) -> tuple[str, str | None]:
    return _source_key(table["source"], base_dir)


def _selector_index(selector: Any, columns: list[dict[str, Any]]) -> int | None:
    if selector is None:
        return None
    if isinstance(selector, int) and not isinstance(selector, bool):
        return selector if 0 <= selector < len(columns) else None
    text = str(selector)
    for column in columns:
        if column.get("name") == text:
            return int(column["index"])
    return None


def _read_options(header: str, skip_rows: str | int, delimiter: str) -> dict[str, Any]:
    """Convert visible controls to the stable reader options contract."""
    header_option: str | bool = "auto" if header == "自动" else header == "有标题行"
    delimiter_option = {
        "自动识别": "auto",
        "逗号 ,": ",",
        "制表符 \\t": "\t",
        "分号 ;": ";",
        "空白分隔": "whitespace",
    }.get(delimiter, "auto")
    return {
        "sheet": None,
        "header": header_option,
        "skip_rows": int(skip_rows),
        "delimiter": delimiter_option,
        "encoding": "auto",
        "missing_values": [""],
        "formula_policy": "cached",
    }


def _job_error_text(exc: BaseException) -> str:
    detail = str(exc).strip() or exc.__class__.__name__
    return f"读取来源失败：{detail}"


def _column_token(text: str) -> str | int:
    """All digits are a zero-based column index. Any other token is a column name."""
    token = text.strip()
    if re.fullmatch(r"\d+", token):
        return int(token)
    return token


def build_global_plot(*, scope: str, kind: str, explicit_x: str, explicit_y: str, explicit_error: str) -> dict[str, Any]:
    """Plot request applied to every file and every sheet.

    ``auto`` leaves column choice to each table's own suggestion. Explicit
    tokens are passed through unchanged so a missing column fails that job
    instead of being replaced by the preview's suggestion.
    """
    plot_kind = kind or "auto"
    if plot_kind not in {"auto", "none", "line", "scatter", "line_symbol", "column"}:
        raise ValueError(f"不支持的图形类型：{plot_kind}")
    if scope not in {"explicit", "明确列"}:
        return {
            "kind": plot_kind,
            "x": "auto",
            "y": "auto",
            "y_error": "auto",
            "title": "auto",
            "x_label": "auto",
            "y_label": "auto",
        }
    y_tokens = [part.strip() for part in explicit_y.replace("，", ",").replace(";", ",").split(",")]
    y_tokens = [part for part in y_tokens if part]
    if plot_kind == "none":
        y_values: list[Any] = []
    elif not y_tokens:
        raise ValueError("明确列需要至少一个 Y 列名或从 0 开始的列号。全数字按列号，其余按列名。")
    else:
        y_values = [_column_token(part) for part in y_tokens]
    x_text = explicit_x.strip()
    plot: dict[str, Any] = {
        "kind": plot_kind,
        "x": _column_token(x_text) if x_text else "auto",
        "y": y_values,
        "y_error": {},
        "title": "auto",
        "x_label": "auto",
        "y_label": "auto",
    }
    error_text = explicit_error.strip()
    if error_text:
        if len(y_tokens) != 1:
            raise ValueError("误差列只在明确选择了单个 Y 列时可用。")
        plot["y_error"] = {y_tokens[0]: _column_token(error_text)}
    return plot


def _format_task_status(event: dict[str, Any]) -> str:
    labels = {
        "succeeded": "成功",
        "failed": "失败",
        "cancelled": "已取消",
        "blocked": "已阻塞",
        "skipped": "已跳过",
    }
    text = labels.get(str(event.get("status") or ""), str(event.get("status") or ""))
    if event.get("pdf_status") == "failed":
        text += "（PDF 失败）"
    return text


def _format_task_detail(event: dict[str, Any]) -> str:
    parts: list[str] = []
    error = event.get("error")
    if isinstance(error, dict):
        message = str(error.get("message") or error.get("type") or "").strip()
        if message:
            parts.append(message)
    elif error:
        parts.append(str(error))
    for item in event.get("pdf_errors") or []:
        if not isinstance(item, dict):
            parts.append(str(item))
            continue
        name = item.get("graph") or item.get("table") or item.get("short_name") or "图"
        reason = item.get("message") or "PDF 失败"
        parts.append(f"{name}: {reason}")
    if event.get("pdf_status") == "failed" and not parts:
        parts.append("PDF 失败")
    return "；".join(parts)


def _format_batch_counts(counts: dict[str, Any]) -> str:
    return (
        f"成功 {int(counts.get('succeeded') or 0)}，"
        f"跳过 {int(counts.get('skipped') or 0)}，"
        f"失败 {int(counts.get('failed') or 0)}，"
        f"取消 {int(counts.get('cancelled') or 0)}，"
        f"阻塞 {int(counts.get('blocked') or 0)}，"
        f"PDF 失败 {int(counts.get('pdf_failed') or 0)}"
    )


_REPARSE_POINT = 0x400
_FILE_ATTRIBUTE_DIRECTORY = 0x10


def _path_is_link(path: Path) -> bool:
    try:
        if path.is_symlink():
            return True
    except OSError:
        return True
    probe = getattr(path, "is_junction", None)
    if callable(probe):
        try:
            return bool(probe())
        except OSError:
            return False
    try:
        attributes = getattr(os.lstat(path), "st_file_attributes", 0)
    except OSError:
        return False
    return bool(attributes & _REPARSE_POINT)


def _entry_is_symlink(entry: os.DirEntry[str]) -> bool:
    try:
        return bool(entry.is_symlink())
    except OSError:
        return False


def _entry_is_directory_link(entry: os.DirEntry[str]) -> bool:
    """True for a junction or a directory symlink.

    On Python 3.12+ / Windows a junction is not a symlink:
    ``is_dir(follow_symlinks=False)`` is true and ``is_junction`` is true.
    A directory symlink is the opposite: ``is_symlink`` is true and
    ``is_dir(follow_symlinks=False)`` is false, so it is classified before
    the non-following file and directory checks.

    Python 3.10 and 3.11 have no ``is_junction``. A non-symlink reparse
    point is a directory link only when it also has the directory
    attribute, which is how a junction still looks when links are not
    followed. A reparse file (OneDrive placeholder or recall-on-data
    ``.csv``) has the reparse attribute without the directory attribute,
    so it stays a file and is not passed to ``scandir``.
    """
    probe = getattr(entry, "is_junction", None)
    if callable(probe):
        try:
            if probe():
                return True
        except OSError:
            return False
    elif not _entry_is_symlink(entry):
        try:
            attributes = getattr(entry.stat(follow_symlinks=False), "st_file_attributes", 0)
        except OSError:
            attributes = 0
        if attributes & _REPARSE_POINT:
            return bool(attributes & _FILE_ATTRIBUTE_DIRECTORY)
    if not _entry_is_symlink(entry):
        return False
    try:
        return bool(entry.is_dir(follow_symlinks=True))
    except OSError:
        return False


def _iter_supported_files(root: Path) -> tuple[list[Path], list[str]]:
    """List supported files under a folder, yielding the GIL between directory reads.

    File symlinks are included when their own suffix is supported. Directory
    symlinks and junctions are followed. ``seen_dirs`` stores each resolved
    target, so a link back to an ancestor is skipped instead of looping.
    A link that cannot be resolved or listed is returned in the rejection
    list. ``discover_files`` uses ``Path.rglob``, which also descends into
    junctions, but it does not keep this set: a junction cycle does not
    finish there.
    """
    found: list[Path] = []
    rejected: list[str] = []
    stack = [root]
    seen_dirs: set[str] = set()
    visited = 0
    while stack:
        current = stack.pop()
        try:
            dir_key = _normalized_path(current)
        except OSError as exc:
            if _path_is_link(Path(current)):
                rejected.append(f"{current}（无法解析目录链接：{exc}）")
            continue
        if dir_key in seen_dirs:
            continue
        seen_dirs.add(dir_key)
        try:
            children = list(os.scandir(current))
        except OSError as exc:
            if _path_is_link(Path(current)):
                rejected.append(f"{current}（无法跟随目录链接：{exc}）")
            continue
        visited += 1
        if visited % 20 == 0:
            time.sleep(0)
        for entry in children:
            path = Path(entry.path)
            try:
                if _entry_is_directory_link(entry):
                    stack.append(path)
                    continue
                if _entry_is_symlink(entry):
                    try:
                        is_file = entry.is_file(follow_symlinks=True)
                    except OSError as exc:
                        rejected.append(f"{path}（无法跟随符号链接：{exc}）")
                        continue
                    if not is_file:
                        rejected.append(f"{path}（符号链接目标不存在或不是可导入文件）")
                        continue
                    if path.suffix.casefold() in SUPPORTED_SUFFIXES:
                        found.append(path.resolve())
                    continue
                if entry.is_dir(follow_symlinks=False):
                    stack.append(path)
                elif entry.is_file(follow_symlinks=False) and path.suffix.casefold() in SUPPORTED_SUFFIXES:
                    found.append(path.resolve())
            except OSError as exc:
                rejected.append(f"{path}（{exc}）")
    if not found and not rejected:
        raise ValueError(f"目录中没有可导入的数据文件：{root}")
    unique: dict[Path, None] = {}
    for path in found:
        unique[path] = None
    return sorted(unique, key=_natural_key), rejected


def _collect_supported(paths: list[str | Path]) -> tuple[list[Path], list[str]]:
    """Expand dropped folders, filter supported data files, and deduplicate."""
    found: list[Path] = []
    rejected: list[str] = []
    seen: set[str] = set()
    for raw_path in paths:
        path = Path(raw_path).expanduser()
        if not path.exists():
            raise FileNotFoundError(f"数据源不存在：{path}")
        if path.is_file() and path.suffix.lower() not in SUPPORTED_SUFFIXES:
            rejected.append(str(path))
            continue
        try:
            if path.is_dir():
                candidates, notes = _iter_supported_files(path.resolve())
                rejected.extend(notes)
            else:
                candidates = discover_files([path.resolve()])
        except ValueError as exc:
            if path.is_dir():
                rejected.append(f"{path}（{exc}）")
                continue
            raise
        for candidate in candidates:
            key = _normalized_path(candidate)
            if key not in seen:
                seen.add(key)
                found.append(candidate)
    return found, rejected


def _resolve_plan_path(raw_path: str | Path, base_dir: Path | None) -> Path:
    path = Path(raw_path).expanduser()
    if not path.is_absolute() and base_dir is not None:
        path = base_dir / path
    return path.resolve(strict=False)


def _default_output_for(paths: list[Path], fmt: str) -> Path:
    """Choose an output beside the first source without replacing an input workbook."""
    suffix = _FORMAT_SUFFIX[fmt]
    first = paths[0] if paths else Path.cwd() / "data_import"
    candidate = first.parent / f"{first.stem}{suffix}"
    input_keys = {_normalized_path(path) for path in paths}
    if _normalized_path(candidate) in input_keys:
        candidate = first.parent / f"{first.stem}_imported{suffix}"
        serial = 2
        while _normalized_path(candidate) in input_keys:
            candidate = first.parent / f"{first.stem}_imported_{serial}{suffix}"
            serial += 1
    return candidate


def _validate_plan_save_target(plan: dict[str, Any], target: Path, base_dir: Path | None) -> None:
    """Reject plan paths that could overwrite an input or planned output."""
    target_key = _normalized_path(target)
    for item in plan.get("tables", []):
        source = item.get("source", {})
        source_path = _resolve_plan_path(source.get("path", ""), base_dir)
        if target_key == _normalized_path(source_path):
            raise ValueError(f"计划文件不能覆盖输入数据源：{source_path}")
    output = plan.get("output", {})
    if output.get("path"):
        output_path = _resolve_plan_path(output["path"], base_dir)
        if target_key == _normalized_path(output_path):
            raise ValueError(f"计划文件不能覆盖计划输出文件：{output_path}")


class GeneralDataApp:
    """Inspect, configure, and import tabular data using a responsive Tk UI."""

    def __init__(self, initial_files: list[Path] | None = None) -> None:
        self.root = tk.Tk()
        self.root.title("数据 → Origin 工程")
        self.root.minsize(1100, 860)

        self._paths: list[Path] = []
        self._path_keys: list[str] = []
        self._table_keys: list[str] = []
        self._key_cache: dict[str, str] = {}
        self._tables: list[dict[str, Any]] = []
        self._tables_by_iid: dict[str, dict[str, Any]] = {}
        self._plots: dict[tuple[str, str | None], dict[str, Any]] = {}
        self._column_labels: dict[tuple[str, str | None], dict[str, Any]] = {}
        self._loaded_plan: dict[str, Any] | None = None
        self._plan_base_dir: Path | None = None
        self._current_key: tuple[str, str | None] | None = None
        self._current_table: dict[str, Any] | None = None
        self._thread: threading.Thread | None = None
        self._thread_queue: queue.Queue[dict[str, Any]] = queue.Queue()
        self._thread_kind: str | None = None
        self._import_process = None
        self._import_recv = None
        self._import_cancel = None
        self._import_kind: str | None = None
        self._import_result: dict[str, Any] | None = None
        self._import_exit_waits = 0
        self._import_eof = False
        self._saved_request: dict[str, Any] | None = None
        self._after_list = None
        self._task_iids: dict[str, str] = {}
        self._task_seq = 0
        self._batch_terminal = False
        self._summaries = SummaryStore()
        self._source_errors: dict[str, str] = {}
        self._fingerprints: dict[tuple[str, str | None], tuple[Any, str]] = {}
        self._iids_by_path: dict[str, list[str]] = {}
        self._pending_outcomes: deque[tuple[dict[str, Any], str]] = deque()
        self._pending_listed: deque[dict[str, Any]] = deque()
        self._suppress_preview = False
        self._job_terminal = False
        self._job_error: str | None = None
        self._done_message: dict[str, Any] | None = None
        self._poll_confirms = 0
        self._iid_seq = 0
        self._closed = False
        self._poll_after = None
        self._hook_after = None
        self._busy = False
        self._interactive: list[tk.Widget] = []
        self._widget_states: dict[tk.Widget, str] = {}

        self.header_var = tk.StringVar(value="自动")
        self.skip_rows_var = tk.StringVar(value="0")
        self.delimiter_var = tk.StringVar(value="自动识别")
        self.plot_kind_var = tk.StringVar(value="line")
        self.x_choice_var = tk.StringVar(value="行号")
        self.error_choice_var = tk.StringVar(value="无")
        self.title_var = tk.StringVar()
        self.x_label_var = tk.StringVar()
        self.y_label_var = tk.StringVar()
        self.format_var = tk.StringVar(value="opju")
        self.keep_open_var = tk.BooleanVar(value=False)
        self.layout_var = tk.StringVar(value="per_file")
        self.pdf_var = tk.BooleanVar(value=False)
        self.batch_overwrite_var = tk.BooleanVar(value=False)
        self.plot_scope_var = tk.StringVar(value="自动（按每张表建议）")
        self.batch_kind_var = tk.StringVar(value="auto")
        self.explicit_x_var = tk.StringVar()
        self.explicit_y_var = tk.StringVar()
        self.explicit_error_var = tk.StringVar()
        self.batch_name_var = tk.StringVar()
        self.status_var = tk.StringVar(
            value="默认每个来源一个文件。添加后只发现路径；预览选中的来源。合并模式才会逐个检查全部来源。"
        )

        self._build()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._guard_destroy()
        self._hook_after = self.root.after(200, self._hook_drop)
        self._schedule_drive()
        if initial_files:
            self._add_paths(initial_files, announce=False)

    def run(self) -> None:
        self.root.mainloop()

    def _track(self, widget: tk.Widget) -> tk.Widget:
        self._interactive.append(widget)
        try:
            self._widget_states[widget] = str(widget.cget("state")) or "normal"
        except tk.TclError:
            self._widget_states[widget] = "normal"
        return widget

    def _build(self) -> None:
        outer = ttk.Frame(self.root, padding=10)
        outer.pack(fill=tk.BOTH, expand=True)

        top = ttk.Frame(outer)
        top.pack(fill=tk.X)
        ttk.Label(top, text="支持文本表格、Excel、JSON 和 JSONL；可把文件或文件夹拖到窗口。").pack(side=tk.LEFT, fill=tk.X, expand=True)
        for label, command in (
            ("添加文件", self.add_files),
            ("添加文件夹", self.add_folder),
            ("移除数据源", self.remove_selected_sources),
        ):
            self._track(ttk.Button(top, text=label, command=command)).pack(side=tk.LEFT, padx=(6, 0))

        source_frame = ttk.LabelFrame(outer, text="数据表和工作表", padding=5)
        source_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        self.table_tree = ttk.Treeview(
            source_frame,
            columns=("source", "sheet", "rows", "columns", "state"),
            show="headings",
            height=6,
            selectmode="browse",
        )
        for column, heading, width, anchor in (
            ("source", "文件", 360, tk.W),
            ("sheet", "工作表", 140, tk.W),
            ("rows", "行数", 70, tk.E),
            ("columns", "列数", 70, tk.E),
            ("state", "状态", 280, tk.W),
        ):
            self.table_tree.heading(column, text=heading)
            self.table_tree.column(column, width=width, anchor=anchor, stretch=column == "source")
        source_scroll = ttk.Scrollbar(source_frame, orient=tk.VERTICAL, command=self.table_tree.yview)
        self.table_tree.configure(yscrollcommand=source_scroll.set)
        self.table_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        source_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.table_tree.bind("<<TreeviewSelect>>", self._on_table_selection)
        self.table_tree.bind("<Double-1>", lambda _event: self.preview_selected())

        read_frame = ttk.LabelFrame(outer, text="读取选项", padding=(7, 5))
        read_frame.pack(fill=tk.X, pady=(7, 0))
        ttk.Label(read_frame, text="标题行").pack(side=tk.LEFT)
        self._track(ttk.Combobox(read_frame, state="readonly", width=10, textvariable=self.header_var, values=("自动", "有标题行", "无标题行"))).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Label(read_frame, text="跳过开头行").pack(side=tk.LEFT)
        self._track(ttk.Spinbox(read_frame, from_=0, to=100000, width=8, textvariable=self.skip_rows_var)).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Label(read_frame, text="分隔符").pack(side=tk.LEFT)
        self._track(ttk.Combobox(read_frame, state="readonly", width=14, textvariable=self.delimiter_var, values=_DELIMITER_OPTIONS)).pack(side=tk.LEFT, padx=(4, 10))
        self._track(ttk.Button(read_frame, text="重新读取", command=self.reload_inputs)).pack(side=tk.LEFT)
        ttk.Label(read_frame, text="修改这些设置后重新读取，计划会保存实际读取选项。", foreground="#555555").pack(side=tk.LEFT, padx=(12, 0))

        preview_box = ttk.LabelFrame(outer, text="列信息与数据预览", padding=5)
        preview_box.pack(fill=tk.BOTH, expand=True, pady=(7, 0))
        preview_panes = ttk.Panedwindow(preview_box, orient=tk.HORIZONTAL)
        preview_panes.pack(fill=tk.BOTH, expand=True)
        meta_frame = ttk.Frame(preview_panes)
        sample_frame = ttk.Frame(preview_panes)
        preview_panes.add(meta_frame, weight=1)
        preview_panes.add(sample_frame, weight=3)
        self.column_tree = ttk.Treeview(meta_frame, columns=("index", "name", "kind", "unit", "missing"), show="headings", height=7)
        for column, heading, width in (
            ("index", "列", 40), ("name", "名称", 150), ("kind", "类型", 78), ("unit", "单位", 65), ("missing", "缺失", 65),
        ):
            self.column_tree.heading(column, text=heading)
            self.column_tree.column(column, width=width, stretch=column == "name")
        self.column_tree.pack(fill=tk.BOTH, expand=True)
        self.sample_tree = ttk.Treeview(sample_frame, show="headings", height=7)
        sample_y = ttk.Scrollbar(sample_frame, orient=tk.VERTICAL, command=self.sample_tree.yview)
        sample_x = ttk.Scrollbar(sample_frame, orient=tk.HORIZONTAL, command=self.sample_tree.xview)
        self.sample_tree.configure(yscrollcommand=sample_y.set, xscrollcommand=sample_x.set)
        self.sample_tree.grid(row=0, column=0, sticky="nsew")
        sample_y.grid(row=0, column=1, sticky="ns")
        sample_x.grid(row=1, column=0, sticky="ew")
        sample_frame.rowconfigure(0, weight=1)
        sample_frame.columnconfigure(0, weight=1)

        plot_box = ttk.LabelFrame(outer, text="当前表的绘图设置（仅“合并为一个工程”时写入计划）", padding=(7, 5))
        plot_box.pack(fill=tk.X, pady=(7, 0))
        row = ttk.Frame(plot_box)
        row.pack(fill=tk.X)
        ttk.Label(row, text="图形").pack(side=tk.LEFT)
        self._track(ttk.Combobox(row, state="readonly", width=14, textvariable=self.plot_kind_var, values=("none", "line", "scatter", "line_symbol", "column"))).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Label(row, text="X 列").pack(side=tk.LEFT)
        self.x_combo = self._track(ttk.Combobox(row, state="readonly", width=32, textvariable=self.x_choice_var, values=("行号",)))
        self.x_combo.pack(side=tk.LEFT, padx=(4, 12))
        ttk.Label(row, text="Y 列（可多选）").pack(side=tk.LEFT)
        ttk.Label(row, text="误差列（单 Y）").pack(side=tk.LEFT, padx=(14, 0))
        self.error_combo = self._track(ttk.Combobox(row, state="readonly", width=26, textvariable=self.error_choice_var, values=("无",)))
        self.error_combo.pack(side=tk.LEFT, padx=(4, 0))
        y_frame = ttk.Frame(plot_box)
        y_frame.pack(fill=tk.X, pady=(4, 0))
        self.y_list = tk.Listbox(y_frame, selectmode=tk.MULTIPLE, height=4, exportselection=False)
        self.y_list.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.y_list.bind("<<ListboxSelect>>", self._update_error_state)
        y_scroll = ttk.Scrollbar(y_frame, orient=tk.VERTICAL, command=self.y_list.yview)
        y_scroll.pack(side=tk.LEFT, fill=tk.Y)
        self.y_list.configure(yscrollcommand=y_scroll.set)
        label_row = ttk.Frame(plot_box)
        label_row.pack(fill=tk.X, pady=(4, 0))
        for label, var, width in (("图题", self.title_var, 30), ("X 轴标题", self.x_label_var, 25), ("Y 轴标题", self.y_label_var, 25)):
            ttk.Label(label_row, text=label).pack(side=tk.LEFT)
            self._track(ttk.Entry(label_row, textvariable=var, width=width)).pack(side=tk.LEFT, padx=(4, 12))

        batch_box = ttk.LabelFrame(outer, text="批量绘图与输出", padding=(7, 5))
        batch_box.pack(fill=tk.X, pady=(7, 0))
        scope = ttk.Label(
            batch_box,
            text="生效范围：每一个文件、每一张表。自动按每张表自己的建议；明确列对不上时该文件失败，不会改用预览里的其他列。全数字按列号。",
            wraplength=1040,
            justify=tk.LEFT,
        )
        scope.pack(fill=tk.X)
        batch_row = ttk.Frame(batch_box)
        batch_row.pack(fill=tk.X, pady=(4, 0))
        ttk.Label(batch_row, text="列选择").pack(side=tk.LEFT)
        self.plot_scope_combo = self._track(ttk.Combobox(
            batch_row, state="readonly", width=22, textvariable=self.plot_scope_var,
            values=("自动（按每张表建议）", "明确列"),
        ))
        self.plot_scope_combo.pack(side=tk.LEFT, padx=(4, 12))
        self.plot_scope_var.trace_add("write", lambda *_args: self._sync_batch_controls())
        ttk.Label(batch_row, text="图形").pack(side=tk.LEFT)
        self.batch_kind_combo = self._track(ttk.Combobox(
            batch_row, state="readonly", width=12, textvariable=self.batch_kind_var,
            values=("auto", "none", "line", "scatter", "line_symbol", "column"),
        ))
        self.batch_kind_combo.pack(side=tk.LEFT, padx=(4, 12))
        ttk.Label(batch_row, text="批次名称").pack(side=tk.LEFT)
        self.batch_name_entry = self._track(ttk.Entry(batch_row, textvariable=self.batch_name_var, width=24))
        self.batch_name_entry.pack(side=tk.LEFT, padx=(4, 0))
        explicit_row = ttk.Frame(batch_box)
        explicit_row.pack(fill=tk.X, pady=(4, 0))
        ttk.Label(explicit_row, text="X").pack(side=tk.LEFT)
        self.explicit_x_entry = self._track(ttk.Entry(explicit_row, textvariable=self.explicit_x_var, width=18))
        self.explicit_x_entry.pack(side=tk.LEFT, padx=(4, 10))
        ttk.Label(explicit_row, text="Y（逗号分隔）").pack(side=tk.LEFT)
        self.explicit_y_entry = self._track(ttk.Entry(explicit_row, textvariable=self.explicit_y_var, width=28))
        self.explicit_y_entry.pack(side=tk.LEFT, padx=(4, 10))
        ttk.Label(explicit_row, text="误差列").pack(side=tk.LEFT)
        self.explicit_error_entry = self._track(ttk.Entry(explicit_row, textvariable=self.explicit_error_var, width=18))
        self.explicit_error_entry.pack(side=tk.LEFT, padx=(4, 0))

        bottom = ttk.Frame(outer)
        bottom.pack(fill=tk.X, pady=(8, 0))
        format_frame = ttk.Frame(bottom)
        format_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(format_frame, text="方式").pack(side=tk.LEFT)
        self._track(ttk.Radiobutton(
            format_frame, text="每个来源一个文件", value="per_file", variable=self.layout_var, command=self._layout_changed,
        )).pack(side=tk.LEFT, padx=(6, 0))
        self._track(ttk.Radiobutton(
            format_frame, text="合并为一个工程", value="unified", variable=self.layout_var, command=self._layout_changed,
        )).pack(side=tk.LEFT, padx=(6, 10))
        self._track(ttk.Radiobutton(format_frame, text="OPJU", value="opju", variable=self.format_var, command=self._format_changed)).pack(side=tk.LEFT)
        self._track(ttk.Radiobutton(format_frame, text="XLSX", value="xlsx", variable=self.format_var, command=self._format_changed)).pack(side=tk.LEFT, padx=(6, 0))
        self.keep_open_check = self._track(ttk.Checkbutton(format_frame, text="完成后保持打开", variable=self.keep_open_var))
        self.keep_open_check.pack(side=tk.LEFT, padx=(8, 0))
        self.pdf_check = self._track(ttk.Checkbutton(format_frame, text="每图 PDF", variable=self.pdf_var))
        self.pdf_check.pack(side=tk.LEFT, padx=(8, 0))
        self.overwrite_check = self._track(ttk.Checkbutton(format_frame, text="覆盖已有输出", variable=self.batch_overwrite_var))
        self.overwrite_check.pack(side=tk.LEFT, padx=(8, 0))

        actions = ttk.Frame(outer)
        actions.pack(fill=tk.X, pady=(6, 0))
        self._track(ttk.Button(actions, text="预览选中", command=self.preview_selected)).pack(side=tk.LEFT, padx=(0, 5))
        self._track(ttk.Button(actions, text="加载计划 JSON", command=self.load_plan)).pack(side=tk.LEFT, padx=(0, 5))
        self._track(ttk.Button(actions, text="保存计划 JSON", command=self.save_plan)).pack(side=tk.LEFT, padx=(0, 5))
        self._track(ttk.Button(actions, text="继续批次…", command=self.continue_batch)).pack(side=tk.LEFT, padx=(0, 5))
        self._track(ttk.Button(actions, text="重试失败项…", command=self.retry_failed_batch)).pack(side=tk.LEFT, padx=(0, 5))
        self._track(ttk.Button(actions, text="批量谱线模式", command=self.open_spectra_gui)).pack(side=tk.LEFT, padx=(0, 5))
        self.export_button = self._track(ttk.Button(actions, text="创建文件…", command=self.export_data))
        self.export_button.pack(side=tk.LEFT)
        self.stop_button = ttk.Button(actions, text="停止后续", command=self.stop_later_tasks)
        self.stop_button.pack(side=tk.LEFT, padx=(8, 0))

        task_frame = ttk.LabelFrame(outer, text="批量任务", padding=5)
        task_frame.pack(fill=tk.BOTH, expand=True, pady=(7, 0))
        self.task_tree = ttk.Treeview(
            task_frame,
            columns=("source", "status", "detail"),
            show="headings",
            height=5,
        )
        for column, heading, width in (
            ("source", "来源", 420),
            ("status", "状态", 140),
            ("detail", "错误或 PDF", 460),
        ):
            self.task_tree.heading(column, text=heading)
            self.task_tree.column(column, width=width, anchor=tk.W, stretch=column != "status")
        task_scroll = ttk.Scrollbar(task_frame, orient=tk.VERTICAL, command=self.task_tree.yview)
        self.task_tree.configure(yscrollcommand=task_scroll.set)
        self.task_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        task_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self._format_changed()

        ttk.Label(outer, textvariable=self.status_var, wraplength=1060, justify=tk.LEFT).pack(fill=tk.X, pady=(8, 0))

    def _hook_drop(self) -> None:
        self._hook_after = None
        if not self._ui_alive():
            return
        try:
            import spectra_to_origin as sto

            sto.enable_windows_file_drop(self.root, self._on_drop_paths)
        except Exception:
            return

    def _on_drop_paths(self, paths: list[str]) -> None:
        if not self._busy:
            self._add_paths(paths)

    def add_files(self) -> None:
        selected = filedialog.askopenfilenames(
            parent=self.root,
            title="选择数据文件",
            filetypes=(("支持的数据文件", "*.txt *.dat *.xy *.csv *.tsv *.xlsx *.xlsm *.xls *.json *.jsonl *.ndjson"), ("所有文件", "*.*")),
        )
        if selected:
            self._add_paths(list(selected))

    def add_folder(self) -> None:
        folder = filedialog.askdirectory(parent=self.root, title="选择包含数据文件的文件夹", mustexist=True)
        if folder:
            self._add_paths([folder])

    def _add_paths(self, raw_paths: list[str | Path], announce: bool = True) -> None:
        if self._busy:
            return
        options = self._try_read_options()
        if options is None:
            return
        requested = [Path(path) for path in raw_paths]
        existing = list(self._paths)
        list_only = self._per_file_mode()
        self.status_var.set("正在后台扫描文件…" if list_only else "正在后台扫描并检查文件…")

        def scan_task() -> None:
            try:
                found, rejected = _collect_supported(requested)
                existing_keys = set()
                for index, path in enumerate(existing):
                    existing_keys.add(_source_key({"path": path})[0])
                    if index % 25 == 0:
                        time.sleep(0)
                added = []
                added_records = []
                for index, path in enumerate(found):
                    key = _source_key({"path": path})[0]
                    if key not in existing_keys:
                        added.append(path)
                        added_records.append((path, key))
                    if index % 25 == 0:
                        time.sleep(0)
                self._thread_queue.put({
                    "kind": "scan",
                    "added": added_records,
                    "rejected": rejected,
                    "announce": announce,
                    "list_only": list_only,
                })
                if list_only:
                    for index, (path, key) in enumerate(added_records):
                        self._thread_queue.put({
                            "kind": "listed",
                            "path": str(path),
                            "path_key": key,
                        })
                        if index % 25 == 0:
                            time.sleep(0)
                else:
                    self._inspect_paths_worker(added, options, "replace")
                self._thread_queue.put({
                    "kind": "inspect_done",
                    "mode": "add",
                    "added_count": len(added),
                    "announce": announce,
                    "rejected_count": len(rejected),
                    "list_only": list_only,
                })
            except Exception as exc:
                self._thread_queue.put({"kind": "job_error", "error": _job_error_text(exc)})

        self._start_thread(scan_task, "scan")

    def remove_selected_sources(self) -> None:
        selected = self.table_tree.selection()
        if not selected or self._busy:
            return
        self._save_current_plot()
        remove_keys = {_normalized_path(self._tables_by_iid[iid]["source"]["path"]) for iid in selected}
        kept_paths: list[Path] = []
        kept_keys: list[str] = []
        for path, key in zip(self._paths, self._path_keys):
            if key not in remove_keys:
                kept_paths.append(path)
                kept_keys.append(key)
        self._paths = kept_paths
        self._path_keys = kept_keys
        kept_tables: list[dict[str, Any]] = []
        kept_table_keys: list[str] = []
        for table, key in zip(self._tables, self._table_keys):
            if key not in remove_keys:
                kept_tables.append(table)
                kept_table_keys.append(key)
        self._tables = kept_tables
        self._table_keys = kept_table_keys
        self._loaded_plan = None
        self._plan_base_dir = None
        for path_key in remove_keys:
            self._source_errors.pop(path_key, None)
            self._drop_path_memory(path_key)
            for iid in self._iids_by_path.pop(path_key, []):
                self._tables_by_iid.pop(iid, None)
                if self.table_tree.exists(iid):
                    self.table_tree.delete(iid)
        if self.table_tree.get_children():
            self._ensure_selection()
            self.status_var.set("已移除数据源。")
        else:
            self._current_key = None
            self._current_table = None
            self._show_table(None)
            self.status_var.set("已移除数据源。")

    def reload_inputs(self) -> None:
        """Re-read every current file with the visible header, skip, and delimiter.

        This drops a loaded plan. Worksheet filtering from that plan is not
        kept: the visible controls have no sheet, so each workbook is listed
        as one row per worksheet. Plots tied to the plan's sheet options are
        dropped when that read identity no longer matches.
        """
        if not self._paths or self._busy:
            return
        self._save_current_plot()
        options = self._try_read_options()
        if options is None:
            return
        self._loaded_plan = None
        self._plan_base_dir = None
        if self._per_file_mode():
            touched_keys = set(self._table_keys) | set(self._source_errors)
            paths = [path for path, key in zip(self._paths, self._path_keys) if key in touched_keys]
            if not paths:
                self.status_var.set("还没有预览过的来源。批量导出会把已发现的路径交给导出核心，不必先读完全部文件。")
                return
        else:
            paths = list(self._paths)
        self.status_var.set("正在后台重新检查来源…")

        def reload_task() -> None:
            try:
                self._inspect_paths_worker(paths, options, "replace")
                self._thread_queue.put({"kind": "inspect_done", "mode": "reload", "added_count": len(paths), "announce": False})
            except Exception as exc:
                self._thread_queue.put({"kind": "job_error", "error": _job_error_text(exc)})

        self._start_thread(reload_task, "inspect")

    def _try_read_options(self) -> dict[str, Any] | None:
        try:
            options = _read_options(self.header_var.get(), self.skip_rows_var.get(), self.delimiter_var.get())
        except (ValueError, tk.TclError) as exc:
            messagebox.showerror("读取选项无效", f"跳过行数必须是非负整数：{exc}", parent=self.root)
            return None
        if options["skip_rows"] < 0:
            messagebox.showerror("读取选项无效", "跳过行数必须大于或等于 0。", parent=self.root)
            return None
        return options

    def _per_file_mode(self) -> bool:
        return getattr(self, "layout_var", None) is not None and self.layout_var.get() == "per_file"

    def preview_selected(self) -> None:
        """Inspect one selected source. Batch export does not require this for every file."""
        if self._busy or not self._per_file_mode():
            return
        selection = self.table_tree.selection()
        if not selection:
            self.status_var.set("先在列表里选一个来源，再预览。")
            return
        table = self._tables_by_iid.get(selection[0])
        if not table:
            return
        path = Path((table.get("source") or {}).get("path") or "")
        if not str(path):
            return
        options = self._try_read_options()
        if options is None:
            return
        self.status_var.set(f"正在预览：{path.name}")

        def preview_task() -> None:
            try:
                self._inspect_paths_worker([path], options, "replace")
                self._thread_queue.put({
                    "kind": "inspect_done",
                    "mode": "preview",
                    "added_count": 1,
                    "announce": False,
                })
            except Exception as exc:
                self._thread_queue.put({"kind": "job_error", "error": _job_error_text(exc)})

        self._start_thread(preview_task, "preview")

    def _layout_changed(self) -> None:
        self._sync_batch_controls()
        if self._busy or self._per_file_mode() or not self._paths:
            return
        inspected = set(self._table_keys) | set(self._source_errors)
        missing = [path for path, key in zip(self._paths, self._path_keys) if key not in inspected]
        if not missing:
            return
        options = self._try_read_options()
        if options is None:
            return
        self.status_var.set("合并模式会检查尚未预览的来源…")

        def inspect_missing() -> None:
            try:
                self._inspect_paths_worker(missing, options, "replace")
                self._thread_queue.put({
                    "kind": "inspect_done",
                    "mode": "reload",
                    "added_count": len(missing),
                    "announce": False,
                })
            except Exception as exc:
                self._thread_queue.put({"kind": "job_error", "error": _job_error_text(exc)})

        self._start_thread(inspect_missing, "inspect")

    def _inspect_paths_worker(self, paths: list[Path], options: dict[str, Any], update: str) -> None:
        for path in paths:
            outcome = self._summaries.inspect_path(path, options)
            outcome["path_key"] = _source_key({"path": outcome.get("path") or path})[0]
            self._thread_queue.put({"kind": "source", "outcome": outcome, "update": update})
            # Yield so the Tk thread can paint while hundreds of sources are parsed.
            time.sleep(0)

    def _inspect_plan_sources(self, plan: dict[str, Any]) -> None:
        plan_tables = copy.deepcopy(plan["tables"])
        base_dir = self._plan_base_dir
        self.status_var.set("正在后台重新读取导入计划中的文件和工作表…")

        def inspect_task() -> None:
            try:
                # One inspect per path and compatible read options. Sheet is
                # selected afterwards so a multi-sheet workbook is not read
                # once per saved row. The path cache stays a single slot.
                shared: dict[tuple[str, str], dict[str, Any]] = {}
                for plan_table in plan_tables:
                    source = plan_table["source"]
                    options = dict(source.get("options") or {})
                    source_path = _resolve_plan_path(source["path"], base_dir)
                    broad = dict(options)
                    broad["sheet"] = None
                    group_key = (_normalized_path(source_path), options_fingerprint(broad))
                    if group_key not in shared:
                        shared[group_key] = self._summaries.inspect_path(source_path, broad)
                    outcome = self._plan_sheet_outcome(shared[group_key], source_path, options)
                    self._thread_queue.put({"kind": "source", "outcome": outcome, "update": "append"})
                self._thread_queue.put({
                    "kind": "inspect_done",
                    "mode": "load_plan",
                    "added_count": len(plan_tables),
                    "announce": False,
                })
            except Exception as exc:
                self._thread_queue.put({"kind": "job_error", "error": _job_error_text(exc)})

        self._start_thread(inspect_task, "load_plan")

    @staticmethod
    def _summary_sheet(table: dict[str, Any]) -> str | None:
        source = table.get("source") or {}
        options = source.get("options") or {}
        sheet = options.get("sheet", source.get("sheet"))
        return str(sheet) if sheet is not None else None

    def _plan_sheet_outcome(self, broad: dict[str, Any], source_path: Path, options: dict[str, Any]) -> dict[str, Any]:
        """Keep one plan row's sheet from a workbook read that asked for every sheet."""
        fingerprint = options_fingerprint(options)
        sheet = options.get("sheet")
        wanted = str(sheet) if sheet is not None else None
        path = str(broad.get("path") or source_path)
        if not broad.get("ok"):
            failed = dict(broad)
            failed["path"] = path
            failed["requested_fingerprint"] = fingerprint
            return failed
        tables = list(broad.get("tables") or [])
        if wanted is not None:
            tables = [table for table in tables if self._summary_sheet(table) == wanted]
        if not tables:
            return {
                "ok": False,
                "path": path,
                "tables": [],
                "error": f"没有工作表：{wanted}",
                "reused": False,
                "sha256": broad.get("sha256") or "",
                "requested_fingerprint": fingerprint,
            }
        return {
            "ok": True,
            "path": path,
            "tables": tables,
            "error": "",
            "reused": bool(broad.get("reused")),
            "sha256": broad.get("sha256") or "",
            "requested_fingerprint": fingerprint,
        }

    def _ui_alive(self) -> bool:
        root = getattr(self, "root", None)
        exists = getattr(root, "winfo_exists", None)
        if not callable(exists):
            return True
        try:
            return bool(exists())
        except tk.TclError:
            return False

    def _after(self, delay_ms: int, callback) -> None:
        if getattr(self, "_closed", False) or not self._ui_alive():
            return
        try:
            after_id = self.root.after(delay_ms, callback)
        except tk.TclError:
            return
        if hasattr(self, "_poll_after"):
            self._poll_after = after_id

    def _cancel_scheduled_ui(self) -> None:
        for after_id in (self._poll_after, self._hook_after):
            if not after_id:
                continue
            try:
                self.root.after_cancel(after_id)
            except tk.TclError:
                pass
        self._poll_after = None
        self._hook_after = None

    def _guard_destroy(self) -> None:
        """Cancel timers and drop import handles before Tcl tears the window down.

        Direct ``root.destroy()`` is the close path a scan or export can hit.
        Cancelling the import poll must still close the parent pipe. A worker
        that is still writing is not killed or joined here; a reaper closes
        it after the process exits and does not touch Tk. A ``<Destroy>``
        binding also runs for child widgets and is the wrong place to touch timers.
        """
        root = self.root
        original = root.destroy

        def destroy(*args, **kwargs):
            if not self._closed:
                self._closed = True
                self._request_batch_cancel()
                self._cancel_scheduled_ui()
                self._abandon_source_job()
                self._release_import_handles()
            return original(*args, **kwargs)

        root.destroy = destroy

    def _abandon_source_job(self) -> None:
        """Drop source-job state after the window is already gone."""
        self._thread = None
        self._thread_kind = None
        self._busy = False
        self._pending_outcomes.clear()
        if getattr(self, "_pending_listed", None) is not None:
            self._pending_listed.clear()

    def _start_thread(self, target, kind: str) -> None:
        if self._busy:
            return
        self._job_terminal = False
        self._job_error = None
        self._done_message = None
        self._poll_confirms = 0
        self._pending_outcomes.clear()
        if getattr(self, "_pending_listed", None) is not None:
            self._pending_listed.clear()
        self._set_busy(True)
        self._thread_kind = kind
        self._thread = threading.Thread(target=target, name=f"origin-bridge-{kind}", daemon=True)
        self._thread.start()
        self._after(16, self._poll_thread)

    def _poll_thread(self) -> None:
        if not self._ui_alive():
            self._abandon_source_job()
            return
        started = time.monotonic()
        handled = 0
        while handled < _UI_MESSAGE_BUDGET and (time.monotonic() - started) < _UI_FRAME_SECONDS:
            try:
                message = self._thread_queue.get_nowait()
            except queue.Empty:
                break
            handled += 1
            self._dispatch_worker_message(message)
        flushed = self._flush_outcomes(started)
        if flushed and self._paths:
            try:
                if self._thread_kind == "scan" and self._per_file_mode():
                    shown = len(self.table_tree.get_children())
                    self.status_var.set(f"正在列出来源 {shown}/{len(self._paths)}…")
                else:
                    seen = len(self._tables) + len(self._source_errors)
                    self.status_var.set(f"正在检查来源 {seen}/{len(self._paths)}…")
            except tk.TclError:
                self._abandon_source_job()
                return
        thread = self._thread
        alive = thread is not None and thread.is_alive()
        more = bool(self._pending_outcomes) or bool(getattr(self, "_pending_listed", None)) or not self._thread_queue.empty()
        if alive or more:
            self._poll_confirms = 0
            self._after(16, self._poll_thread)
            return
        if thread is not None:
            thread.join(timeout=0.2)
            if thread.is_alive() or not self._thread_queue.empty() or self._pending_outcomes or getattr(self, "_pending_listed", None):
                self._after(16, self._poll_thread)
                return
        if not self._job_terminal and self._poll_confirms < 1:
            self._poll_confirms += 1
            self._after(16, self._poll_thread)
            return
        if not self._job_terminal:
            self._job_error = self._job_error or "后台读取意外结束。"
            self._job_terminal = True
        self._finish_source_job()

    def _dispatch_worker_message(self, message: dict[str, Any]) -> None:
        kind = message.get("kind")
        if kind == "job_error":
            self._job_error = str(message.get("error") or "读取失败")
            self._job_terminal = True
            return
        if kind == "scan":
            self._apply_scan(message)
            return
        if kind == "source":
            outcome = message.get("outcome")
            if isinstance(outcome, dict):
                self._pending_outcomes.append((outcome, str(message.get("update") or "replace")))
            return
        if kind == "listed":
            self._pending_listed.append(message)
            return
        if kind == "inspect_done":
            self._done_message = message
            self._job_terminal = True

    def _apply_scan(self, message: dict[str, Any]) -> None:
        rejected = [str(item) for item in message.get("rejected") or []]
        added = list(message.get("added") or [])
        if rejected:
            messagebox.showwarning(
                "部分来源未加入",
                "以下来源没有加入：\n" + "\n".join(rejected[:8]),
                parent=self.root,
            )
        if added:
            self._save_current_plot()
            self._remember_added(added)
            self._loaded_plan = None
            self._plan_base_dir = None
            return
        if message.get("announce", True) and not rejected:
            self.status_var.set("没有发现新的支持格式文件。")

    def _flush_outcomes(self, started: float) -> int:
        flushed = 0
        listed = getattr(self, "_pending_listed", None)
        if listed is None:
            listed = deque()
            self._pending_listed = listed
        while listed and flushed < _UI_ROW_BUDGET and (time.monotonic() - started) < _UI_FRAME_SECONDS:
            self._insert_listed_row(listed.popleft())
            flushed += 1
        while self._pending_outcomes and flushed < _UI_ROW_BUDGET and (time.monotonic() - started) < _UI_FRAME_SECONDS:
            outcome, update = self._pending_outcomes.popleft()
            self._apply_outcome(outcome, update)
            flushed += 1
        return flushed

    def _cached_path_key(self, path: str | Path) -> str:
        text = os.path.normcase(str(path))
        cached = self._key_cache.get(text)
        if cached is None:
            cached = _source_key({"path": path})[0]
            self._key_cache[text] = cached
            self._key_cache[cached] = cached
        return cached

    def _table_identity(self, source: dict[str, Any]) -> tuple[str, str | None]:
        options = source.get("options") or {}
        sheet = options.get("sheet", source.get("sheet"))
        return self._cached_path_key(source.get("path", "")), str(sheet) if sheet is not None else None

    def _remember_key(self, path: str | Path, key: str) -> None:
        self._key_cache[os.path.normcase(str(path))] = key
        self._key_cache[key] = key

    def _remember_added(self, added: list[Any]) -> None:
        for item in added:
            if isinstance(item, tuple):
                path, key = item
            else:
                path, key = item, self._cached_path_key(item)
            self._paths.append(Path(path))
            self._path_keys.append(key)
            self._remember_key(path, key)

    def _apply_outcome(self, outcome: dict[str, Any], update: str) -> None:
        path = str(outcome.get("path") or "")
        path_key = str(outcome.get("path_key") or self._cached_path_key(path))
        self._remember_key(path, path_key)
        for table in outcome.get("tables") or []:
            source_path = (table.get("source") or {}).get("path")
            if source_path:
                self._remember_key(source_path, path_key)
        if not outcome.get("ok"):
            error = str(outcome.get("error") or "读取失败")
            self._source_errors[path_key] = error
            error_row = self._error_row(path, error)
            if update == "append":
                self._append_row(error_row)
            else:
                self._remove_path_tables(path_key)
                self._drop_path_memory(path_key)
                self._replace_path_rows(path_key, [error_row])
            return
        requested = str(outcome.get("requested_fingerprint") or "")
        fresh_tables = list(outcome.get("tables") or [])
        if update != "append" and outcome.get("reused") and self._same_cached_rows(path_key, fresh_tables, requested):
            self._source_errors.pop(path_key, None)
            return
        if update != "append":
            self._source_errors.pop(path_key, None)
            self._remove_path_tables(path_key)
        for table in fresh_tables:
            source = table.get("source") or {}
            key = self._table_identity(source)
            new_fp = (source.get("sha256"), requested)
            old_fp = self._fingerprints.get(key)
            if old_fp is not None and old_fp != new_fp:
                self._plots.pop(key, None)
                self._column_labels.pop(key, None)
            self._fingerprints[key] = new_fp
            self._plots.setdefault(key, copy.deepcopy(table.get("suggested_plot") or self._default_plot(table)))
        if update == "append":
            for table in fresh_tables:
                self._tables.append(table)
                self._table_keys.append(path_key)
                self._append_row(table)
            return
        self._insert_tables_in_path_order(path_key, fresh_tables)
        self._replace_path_rows(path_key, fresh_tables)

    def _finish_source_job(self) -> None:
        if not self._ui_alive():
            self._abandon_source_job()
            return
        done = self._done_message or {}
        error = self._job_error
        self._thread = None
        self._thread_kind = None
        self._job_error = None
        self._done_message = None
        try:
            self._set_busy(False)
            self._ensure_selection()
        except tk.TclError:
            self._busy = False
            return
        if error:
            try:
                self.status_var.set("读取失败。")
                messagebox.showerror("数据读取失败", error, parent=self.root)
            except tk.TclError:
                self._busy = False
            return
        if done.get("list_only"):
            failed = len(self._source_errors)
            text = f"已发现 {len(self._paths)} 个来源。每个文件会单独导出；预览只检查选中的来源。"
            if failed:
                text += f" {failed} 个已预览来源有错误，其余仍可导出。"
            try:
                self.status_var.set(text)
            except tk.TclError:
                self._busy = False
            return
        if done.get("mode") == "preview":
            try:
                self.status_var.set("预览完成。未预览的来源仍会按路径进入批量导出。")
            except tk.TclError:
                self._busy = False
            return
        if done.get("mode") == "add" and not done.get("added_count") and not self._source_errors:
            return
        failed = len(self._source_errors)
        if done.get("mode") == "load_plan":
            text = f"计划数据已读取，共 {len(self._tables)} 个数据表；导出前会再次验证源文件。"
        else:
            text = f"已读取 {len(self._tables)} 个数据表。"
        if failed:
            text += f" {failed} 个来源读取失败，单工程不会导出不完整计划。"
        elif len(self._tables) <= 3:
            warnings = [str(warning) for table in self._tables for warning in (table.get("warnings") or [])]
            if warnings:
                text += " " + "；".join(warnings[:3])
        try:
            self.status_var.set(text)
        except tk.TclError:
            self._busy = False

    def _same_cached_rows(self, path_key: str, tables: list[dict[str, Any]], requested: str) -> bool:
        iids = self._iids_by_path.get(path_key) or []
        if len(iids) != len(tables) or not iids:
            return False
        for table in tables:
            source = table.get("source") or {}
            key = self._table_identity(source)
            if self._fingerprints.get(key) != (source.get("sha256"), requested):
                return False
            if key not in self._plots:
                return False
        return all(self.table_tree.exists(iid) for iid in iids)

    def _error_row(self, path: str, error: str) -> dict[str, Any]:
        return {
            "source": {"path": path, "sheet": None, "sha256": "", "options": {}},
            "name": Path(path).stem,
            "n_rows": 0,
            "columns": [],
            "warnings": [],
            "error": error,
            "suggested_plot": None,
        }

    def _drop_path_memory(self, path_key: str) -> None:
        for key in [key for key in self._plots if key[0] == path_key]:
            self._plots.pop(key, None)
            self._column_labels.pop(key, None)
            self._fingerprints.pop(key, None)
        for key in [key for key in self._fingerprints if key[0] == path_key]:
            self._fingerprints.pop(key, None)

    def _remove_path_tables(self, path_key: str) -> None:
        self._align_table_keys()
        kept_tables: list[dict[str, Any]] = []
        kept_keys: list[str] = []
        for table, key in zip(self._tables, self._table_keys):
            if key != path_key:
                kept_tables.append(table)
                kept_keys.append(key)
        self._tables = kept_tables
        self._table_keys = kept_keys

    def _align_table_keys(self) -> None:
        if len(self._table_keys) == len(self._tables):
            return
        self._table_keys = [self._cached_path_key((table.get("source") or {}).get("path", "")) for table in self._tables]

    def _insert_tables_in_path_order(self, path_key: str, tables: list[dict[str, Any]]) -> None:
        self._align_table_keys()
        if len(self._path_keys) != len(self._paths):
            self._path_keys = [self._cached_path_key(path) for path in self._paths]
        grouped: dict[str, list[dict[str, Any]]] = {}
        for table, key in zip(self._tables, self._table_keys):
            grouped.setdefault(key, []).append(table)
        grouped[path_key] = list(tables)
        ordered: list[dict[str, Any]] = []
        ordered_keys: list[str] = []
        seen: set[str] = set()
        for key in self._path_keys:
            if key in seen:
                continue
            seen.add(key)
            for table in grouped.get(key, []):
                ordered.append(table)
                ordered_keys.append(key)
        for key, items in grouped.items():
            if key not in seen:
                for table in items:
                    ordered.append(table)
                    ordered_keys.append(key)
        self._tables = ordered
        self._table_keys = ordered_keys

    def _insert_listed_row(self, item: dict[str, Any]) -> None:
        path = str(item.get("path") or "")
        path_key = str(item.get("path_key") or self._cached_path_key(path))
        if path_key in self._iids_by_path or path_key in set(self._table_keys) or path_key in self._source_errors:
            return
        placeholder = {
            "source": {"path": path, "sheet": None, "sha256": "", "options": {}},
            "name": Path(path).stem,
            "n_rows": "",
            "columns": [],
            "warnings": [],
            "pending": True,
        }
        iid = self._insert_row_at(placeholder, tk.END)
        try:
            self.table_tree.see(iid)
        except tk.TclError:
            return

    def _row_values(self, table: dict[str, Any]) -> tuple[Any, ...]:
        source = table.get("source") or {}
        if table.get("error"):
            return (source.get("path", ""), "", "错误", "", table.get("error", ""))
        if table.get("pending"):
            return (source.get("path", ""), "", "", "", "未预览")
        return (
            source.get("path", ""),
            source.get("sheet") or "",
            table.get("n_rows", ""),
            len(table.get("columns") or []),
            "就绪",
        )

    def _next_iid(self) -> str:
        self._iid_seq += 1
        return f"s{self._iid_seq}"

    def _insert_row_at(self, table: dict[str, Any], index: int | str) -> str:
        iid = self._next_iid()
        self._tables_by_iid[iid] = table
        self.table_tree.insert("", index, iid=iid, values=self._row_values(table))
        path_key = self._cached_path_key((table.get("source") or {}).get("path", ""))
        self._iids_by_path.setdefault(path_key, []).append(iid)
        return iid

    def _append_row(self, table: dict[str, Any]) -> None:
        self._insert_row_at(table, tk.END)

    def _replace_path_rows(self, path_key: str, tables: list[dict[str, Any]]) -> None:
        old = self._iids_by_path.get(path_key, [])
        index: int | str = "end"
        if old and self.table_tree.exists(old[0]):
            index = self.table_tree.index(old[0])
        for iid in old:
            self._tables_by_iid.pop(iid, None)
            if self.table_tree.exists(iid):
                self.table_tree.delete(iid)
        new_iids: list[str] = []
        # _insert_row_at appends to _iids_by_path; seed it empty and refill below.
        self._iids_by_path[path_key] = []
        for offset, table in enumerate(tables):
            if index == "end":
                new_iids.append(self._insert_row_at(table, tk.END))
            else:
                new_iids.append(self._insert_row_at(table, int(index) + offset))
        self._iids_by_path[path_key] = new_iids

    def _clear_listed_sources(self) -> None:
        self._tables.clear()
        self._table_keys.clear()
        self._tables_by_iid.clear()
        self._iids_by_path.clear()
        self._pending_outcomes.clear()
        if getattr(self, "_pending_listed", None) is not None:
            self._pending_listed.clear()
        children = self.table_tree.get_children()
        if children:
            self.table_tree.delete(*children)
        self._current_key = None
        self._current_table = None
        self._show_table(None)

    def _ensure_selection(self) -> None:
        children = self.table_tree.get_children()
        if not children:
            self._current_key = None
            self._current_table = None
            self._show_table(None)
            return
        wanted = self._current_key
        chosen = None
        if wanted is not None:
            for iid in children:
                table = self._tables_by_iid.get(iid)
                if table is not None and self._table_identity(table["source"]) == wanted:
                    chosen = iid
                    break
        if chosen is None:
            current = self.table_tree.selection()
            if current and current[0] in self._tables_by_iid:
                chosen = current[0]
            else:
                chosen = children[0]
        self._suppress_preview = True
        try:
            self.table_tree.selection_set(chosen)
            self.table_tree.focus(chosen)
            self._select_table(self._tables_by_iid[chosen])
        finally:
            self._suppress_preview = False

    @staticmethod
    def _default_plot(table: dict[str, Any]) -> dict[str, Any]:
        columns = table.get("columns", [])
        numeric = [int(column["index"]) for column in columns if column.get("kind") == "number"]
        if len(numeric) >= 2:
            x_index, y_indexes, kind = numeric[0], [numeric[1]], "line"
        elif numeric:
            x_index, y_indexes, kind = None, [numeric[0]], "line"
        else:
            x_index, y_indexes, kind = None, [], "none"
        return {
            "kind": kind,
            "x": x_index,
            "y": y_indexes,
            "y_error": {},
            "title": table.get("name", ""),
            "x_label": str(columns[0].get("name", "")) if columns else "",
            "y_label": "",
        }

    def _on_table_selection(self, _event=None) -> None:
        if self._busy:
            return
        self._save_current_plot()
        selection = self.table_tree.selection()
        table = self._tables_by_iid.get(selection[0]) if selection else None
        self._select_table(table)

    def _select_table(self, table: dict[str, Any] | None) -> None:
        self._current_key = self._table_identity(table["source"]) if table else None
        self._current_table = table
        self._show_table(table)

    def _clear_source_controls(self) -> None:
        """Drop the previous table's axes so an error or empty selection cannot keep them."""
        self.sample_tree.configure(columns=())
        self.sample_tree["show"] = "headings"
        self.y_list.delete(0, tk.END)
        self.x_combo.configure(values=("行号",))
        self.error_combo.configure(values=("无",))
        self.plot_kind_var.set("none")
        self.x_choice_var.set("行号")
        self.error_choice_var.set("无")
        self.title_var.set("")
        self.x_label_var.set("")
        self.y_label_var.set("")
        try:
            self.error_combo.configure(state="disabled")
        except tk.TclError:
            pass

    def _show_table(self, table: dict[str, Any] | None) -> None:
        self.column_tree.delete(*self.column_tree.get_children())
        self.sample_tree.delete(*self.sample_tree.get_children())
        if table is not None and table.get("pending"):
            self._clear_source_controls()
            self.status_var.set("此来源尚未预览。批量导出仍会包含它；预览不会把这张表的列建议写成整批的明确列。")
            return
        if table is not None and table.get("error"):
            self._clear_source_controls()
            self.status_var.set(f"来源读取失败：{table['error']}")
            return
        if table is None:
            self._clear_source_controls()
            return

        columns = table.get("columns", [])
        for col in columns:
            self.column_tree.insert("", tk.END, values=(col.get("index", ""), col.get("name", ""), col.get("kind", ""), col.get("unit", ""), col.get("missing", "")))

        sample_ids = [f"c{index}" for index in range(len(columns))]
        self.sample_tree.configure(columns=sample_ids)
        for index, col in enumerate(columns):
            label = str(col.get("name", f"Column {index + 1}"))
            self.sample_tree.heading(sample_ids[index], text=label)
            self.sample_tree.column(sample_ids[index], width=125, minwidth=70, stretch=False)
        samples = [col.get("sample") or [] for col in columns]
        row_count = min(12, max((len(values) for values in samples), default=0))
        for row_index in range(row_count):
            values = [str(sample[row_index]) if row_index < len(sample) and sample[row_index] is not None else "" for sample in samples]
            self.sample_tree.insert("", tk.END, values=values)

        labels = [f"{col.get('index')}: {col.get('name', '')} ({col.get('kind', '')})" for col in columns]
        self.x_combo.configure(values=("行号", *labels))
        numeric_columns = [col for col in columns if str(col.get("kind", "")).lower() in {"number", "numeric", "float", "integer", "int"}]
        numeric = [f"{col.get('index')}: {col.get('name', '')} ({col.get('kind', '')})" for col in numeric_columns]
        self.y_list.delete(0, tk.END)
        for value in numeric:
            self.y_list.insert(tk.END, value)
        error_labels = [f"{col.get('index')}: {col.get('name', '')} ({col.get('kind', '')})" for col in numeric_columns]
        self.error_combo.configure(values=("无", *error_labels))
        self._load_plot_controls(table, numeric)

    def _load_plot_controls(self, table: dict[str, Any], numeric: list[str]) -> None:
        key = self._table_identity(table["source"])
        plot = copy.deepcopy(self._plots.get(key) or table.get("suggested_plot") or self._default_plot(table))
        columns = table.get("columns", [])
        self.plot_kind_var.set(str(plot.get("kind", "line")))
        x_index = _selector_index(plot.get("x"), columns)
        self.x_choice_var.set("行号" if x_index is None else self._selector_label(x_index, columns))
        ys = plot.get("y", []) or []
        if not isinstance(ys, list):
            ys = [ys]
        self.y_list.selection_clear(0, tk.END)
        numeric_indexes = [int(label.split(":", 1)[0]) for label in numeric]
        for selector in ys:
            index = _selector_index(selector, columns)
            if index in numeric_indexes:
                self.y_list.selection_set(numeric_indexes.index(index))
        error_selector: Any = None
        error_mapping = plot.get("y_error", {}) or {}
        if len(ys) == 1 and isinstance(error_mapping, dict):
            y_index = _selector_index(ys[0], columns)
            if y_index is not None:
                error_selector = error_mapping.get(str(columns[y_index].get("name", "")))
        error_index = _selector_index(error_selector, columns)
        self.error_choice_var.set("无" if error_index is None else self._selector_label(error_index, columns))
        self.title_var.set(str(plot.get("title", table.get("name", ""))))
        self.x_label_var.set(str(plot.get("x_label", "")))
        self.y_label_var.set(str(plot.get("y_label", "")))
        self._update_error_state()

    @staticmethod
    def _selector_label(index: int, columns: list[dict[str, Any]]) -> str:
        if 0 <= index < len(columns):
            col = columns[index]
            return f"{col.get('index', index)}: {col.get('name', '')} ({col.get('kind', '')})"
        return "行号"

    @staticmethod
    def _display_index(value: str) -> int | None:
        if value == "行号" or ":" not in value:
            return None
        try:
            return int(value.split(":", 1)[0])
        except ValueError:
            return None

    def _save_current_plot(self) -> None:
        key = self._current_key
        if key is None or (self._current_table or {}).get("error"):
            return
        selected_y = [self.y_list.get(index) for index in self.y_list.curselection()]
        y_indexes = [index for value in selected_y if (index := self._display_index(value)) is not None]
        error_index = self._display_index(self.error_choice_var.get())
        kind = self.plot_kind_var.get() or "none"
        columns = (self._current_table or {}).get("columns", [])
        x_index = None if kind == "none" else self._display_index(self.x_choice_var.get())
        if kind == "none":
            y_indexes = []
        y_error = {}
        if len(y_indexes) == 1 and error_index is not None and y_indexes[0] < len(columns):
            y_error[str(columns[y_indexes[0]].get("name", y_indexes[0]))] = error_index
        plot: dict[str, Any] = {
            "kind": kind,
            "x": x_index,
            "y": y_indexes,
            "y_error": y_error,
            "title": self.title_var.get(),
            "x_label": self.x_label_var.get(),
            "y_label": self.y_label_var.get(),
        }
        self._plots[key] = plot

    def _update_error_state(self, _event=None) -> None:
        single_y = len(self.y_list.curselection()) == 1
        if not single_y:
            self.error_choice_var.set("无")
        try:
            self.error_combo.configure(state="readonly" if single_y else "disabled")
        except tk.TclError:
            pass

    def _format_changed(self) -> None:
        is_origin = self.format_var.get() == "opju"
        if not is_origin:
            self.keep_open_var.set(False)
            self.pdf_var.set(False)
        state = "normal" if is_origin else "disabled"
        try:
            self.keep_open_check.configure(state=state)
            self._widget_states[self.keep_open_check] = state
        except (AttributeError, tk.TclError):
            pass
        self._sync_batch_controls()

    def _sync_batch_controls(self) -> None:
        per_file = self._per_file_mode()
        explicit = per_file and self.plot_scope_var.get() in {"explicit", "明确列"}
        origin = self.format_var.get() == "opju"
        pairs = (
            (getattr(self, "explicit_x_entry", None), explicit),
            (getattr(self, "explicit_y_entry", None), explicit),
            (getattr(self, "explicit_error_entry", None), explicit),
            (getattr(self, "batch_kind_combo", None), per_file),
            (getattr(self, "plot_scope_combo", None), per_file),
            (getattr(self, "batch_name_entry", None), per_file),
            (getattr(self, "pdf_check", None), per_file and origin),
            (getattr(self, "overwrite_check", None), per_file),
        )
        for widget, enabled in pairs:
            if widget is None:
                continue
            state = "normal" if enabled else "disabled"
            try:
                if widget in (getattr(self, "plot_scope_combo", None), getattr(self, "batch_kind_combo", None)):
                    state = "readonly" if enabled else "disabled"
                widget.configure(state=state)
                self._widget_states[widget] = state
            except tk.TclError:
                continue
        if not (per_file and origin):
            self.pdf_var.set(False)
        self._sync_stop_button()

    def _sync_stop_button(self) -> None:
        button = getattr(self, "stop_button", None)
        if button is None:
            return
        enabled = bool(self._busy and getattr(self, "_import_kind", None) == "batch" and self._import_cancel is not None)
        try:
            button.configure(state="normal" if enabled else "disabled")
        except tk.TclError:
            pass

    def _reject_incomplete_sources(self) -> None:
        """Single-project plans refuse failed or not-yet-inspected sources.

        A later batch phase must not treat "every inspect succeeded" as given.
        ``batch_source_snapshot`` exposes per-source errors separately from summaries.
        """
        errors = getattr(self, "_source_errors", None) or {}
        if errors:
            path, detail = next(iter(errors.items()))
            raise ValueError(f"有数据源读取失败，不能保存或导出不完整的单工程计划：{path}：{detail}")
        paths = getattr(self, "_paths", None) or []
        tables = [table for table in (getattr(self, "_tables", None) or []) if not table.get("error")]
        inspected = {_source_key(table["source"])[0] for table in tables if table.get("source")}
        missing = {_normalized_path(path) for path in paths} - inspected
        if not missing:
            return
        if getattr(self, "_loaded_plan", None) is not None:
            raise ValueError("有数据源尚未完成读取或读取失败，不能保存或导出不完整的单工程计划。")
        raise ValueError("有数据源尚未完成读取。请点击“重新读取”并等待完成后再保存计划或导出。")

    def _effective_plan(self, output: Path, fmt: str, overwrite: bool) -> dict[str, Any]:
        self._save_current_plot()
        self._reject_incomplete_sources()
        expected_suffix = _FORMAT_SUFFIX.get(fmt)
        if expected_suffix is None:
            raise ValueError(f"不支持的输出格式：{fmt}")
        if output.suffix.lower() != expected_suffix:
            output = output.with_suffix(expected_suffix)
        if self._loaded_plan is not None:
            plan = copy.deepcopy(self._loaded_plan)
            active_paths = {_normalized_path(path) for path in self._paths}
            plan["tables"] = [
                table for table in plan["tables"]
                if _normalized_path(_resolve_plan_path(table["source"]["path"], self._plan_base_dir)) in active_paths
            ]
        else:
            if not self._paths:
                raise ValueError("请先添加至少一个数据文件。")
            plan_tables = []
            for table in self._tables:
                if table.get("error"):
                    continue
                source = table["source"]
                key = _source_key(source)
                if getattr(self, "_key_cache", None) is not None:
                    alternate = self._table_identity(source)
                    if key not in self._plots and alternate in self._plots:
                        key = alternate
                plan_tables.append({
                    "source": {field: source[field] for field in ("path", "sha256", "options")},
                    "name": table["name"],
                    "plot": copy.deepcopy(self._plots.get(key) or table.get("suggested_plot") or self._default_plot(table)),
                    "column_labels": copy.deepcopy(self._column_labels.get(key, {})),
                })
            plan = {
                "schema_version": PLAN_SCHEMA_VERSION,
                "output": {
                    "path": str(output.resolve()),
                    "format": fmt,
                    "overwrite": bool(overwrite),
                    "keep_open": bool(self.keep_open_var.get()) if fmt == "opju" else False,
                },
                "tables": plan_tables,
            }
        if not isinstance(plan, dict) or plan.get("schema_version") != PLAN_SCHEMA_VERSION:
            raise ValueError("规划模块返回了不支持的导入计划版本。")
        plan_output = plan.get("output")
        if not isinstance(plan_output, dict):
            raise ValueError("导入计划缺少有效的 output 对象。")
        plan_output.update({
            "path": str(output.resolve()),
            "format": fmt,
            "overwrite": bool(overwrite),
            "keep_open": bool(self.keep_open_var.get()) if fmt == "opju" else False,
        })
        if not plan.get("tables"):
            raise ValueError("导入计划没有可导入的数据表。")
        for item in plan["tables"]:
            key = _plan_table_key(item, self._plan_base_dir if self._loaded_plan is not None else None)
            if key in self._plots:
                item["plot"] = copy.deepcopy(self._plots[key])
            item.setdefault("column_labels", {})
        return plan

    def _default_output_path(self, fmt: str) -> Path:
        return _default_output_for(self._paths, fmt)

    def _current_batch_plot(self) -> dict[str, Any]:
        return build_global_plot(
            scope=self.plot_scope_var.get(),
            kind=self.batch_kind_var.get(),
            explicit_x=self.explicit_x_var.get(),
            explicit_y=self.explicit_y_var.get(),
            explicit_error=self.explicit_error_var.get(),
        )

    def _export_per_file(self) -> None:
        if not self._paths:
            messagebox.showinfo("没有数据", "请先添加数据文件或文件夹。", parent=self.root)
            return
        fmt = self.format_var.get()
        if fmt not in _FORMAT_SUFFIX:
            messagebox.showerror("输出格式无效", "请选择 OPJU 或 XLSX。", parent=self.root)
            return
        if self.pdf_var.get() and fmt != "opju":
            messagebox.showerror("不能导出 PDF", "每图 PDF 只适用于每个来源一个 OPJU。", parent=self.root)
            return
        options = self._try_read_options()
        if options is None:
            return
        try:
            plot = self._current_batch_plot()
        except ValueError as exc:
            messagebox.showerror("绘图设置无效", str(exc), parent=self.root)
            return
        if fmt == "opju" and not self._confirm_origin_idle():
            return
        output_dir = filedialog.askdirectory(parent=self.root, title="选择批量输出目录", mustexist=False)
        if not output_dir:
            return
        try:
            from .batch import build_batch_request

            request = build_batch_request(
                self._paths,
                output_dir,
                layout="per_file",
                format=fmt,
                read_options=options,
                plot=plot,
                overwrite=bool(self.batch_overwrite_var.get()),
                pdf=bool(self.pdf_var.get()) if fmt == "opju" else False,
                keep_open=bool(self.keep_open_var.get()) if fmt == "opju" else False,
                resume=False,
                retry_task_ids=[],
            )
        except Exception as exc:
            messagebox.showerror("无法创建批量请求", str(exc), parent=self.root)
            return
        self._saved_request = request
        self._launch_batch(request, name=self.batch_name_var.get())

    def _confirm_origin_idle(self) -> bool:
        try:
            import spectra_to_origin as sto

            if sto.origin_process_running():
                messagebox.showwarning(
                    "请先关闭 Origin",
                    "检测到 Origin 正在运行。请先保存并关闭当前 Origin 工程，再重新创建文件；程序不会重置当前会话。",
                    parent=self.root,
                )
                return False
        except Exception as exc:
            messagebox.showerror("Origin 状态检查失败", str(exc), parent=self.root)
            return False
        return True

    def _launch_batch(self, request: dict[str, Any], *, name: str) -> None:
        if self._busy:
            return
        try:
            json.dumps(request, ensure_ascii=False)
            save_batch_config(
                config_path_for(request["record_path"]),
                name=name,
                record_path=request["record_path"],
                inputs=list(request["inputs"]),
                request=request,
            )
        except Exception as exc:
            messagebox.showerror("无法保存批次配置", str(exc), parent=self.root)
            self.status_var.set("批次配置没有写入，导出没有开始。")
            return
        self.status_var.set("正在后台批量导出；窗口仍可响应。停止后续会在当前项结束后生效。")
        try:
            self._start_batch_process(request)
        except Exception as exc:
            self._set_busy(False)
            messagebox.showerror("无法启动批量导出", str(exc), parent=self.root)
            self.status_var.set("无法启动后台批量导出。")

    def _start_batch_process(self, request: dict[str, Any]) -> None:
        context = multiprocessing.get_context("spawn")
        recv_conn, send_conn = context.Pipe(duplex=False)
        cancel = context.Event()
        drive_spec = os.environ.pop("ORI_GUI_BATCH_DRIVE", None)
        process = context.Process(target=batch_worker, args=(send_conn, request, cancel))
        try:
            process.start()
        except Exception:
            recv_conn.close()
            send_conn.close()
            raise
        finally:
            if drive_spec is not None:
                os.environ["ORI_GUI_BATCH_DRIVE"] = drive_spec
        send_conn.close()
        self._import_recv = recv_conn
        self._import_process = process
        self._import_cancel = cancel
        self._import_kind = "batch"
        self._import_result = None
        self._import_exit_waits = 0
        self._import_eof = False
        self._batch_terminal = False
        self._clear_task_rows()
        self._set_busy(True)
        self._after(100, self._poll_import)

    def _clear_task_rows(self) -> None:
        self._task_iids.clear()
        children = self.task_tree.get_children()
        if children:
            self.task_tree.delete(*children)

    def _request_batch_cancel(self) -> None:
        cancel = getattr(self, "_import_cancel", None)
        if cancel is None:
            return
        try:
            cancel.set()
        except Exception:
            return

    def stop_later_tasks(self) -> None:
        """Ask the batch to finish the current task and not start later ones.

        This does not interrupt a COM call that has not yet proved which Origin
        process it owns. A stuck startup is not ended from this button.
        """
        if getattr(self, "_import_cancel", None) is None:
            return
        self._request_batch_cancel()
        try:
            self.status_var.set("已请求停止后续任务。当前这项会跑完并保存；尚未开始的会取消。这不能打断卡在启动证明之前的 Origin。")
        except tk.TclError:
            return

    def continue_batch(self) -> None:
        self._choose_resume(retry=False)

    def retry_failed_batch(self) -> None:
        self._choose_resume(retry=True)

    def _choose_resume(self, *, retry: bool) -> None:
        if self._busy or not self._per_file_mode():
            if not self._busy and not self._per_file_mode():
                messagebox.showinfo("只用于每个来源一个文件", "PDF、继续和重试失败项只在“每个来源一个文件”模式下可用。合并工程仍使用原来的单次导出。", parent=self.root)
            return
        selected = filedialog.askopenfilename(
            parent=self.root,
            title="选择批量运行记录",
            filetypes=(("批量记录", "*.json"), ("所有文件", "*.*")),
        )
        if not selected:
            return
        record_path = Path(selected)
        config_path = find_config_for_record(record_path)
        if config_path is None:
            proceed = messagebox.askokcancel(
                "缺少批次配置",
                "这份运行记录没有配套的完整批次配置。核心回执不能还原读取选项和绘图设置。\n\n"
                "请接着选择原来的批次配置。取消后不会按当前窗口的默认值重新导出。",
                parent=self.root,
            )
            if not proceed:
                self.status_var.set("没有批次配置，已取消继续。")
                return
            chosen = filedialog.askopenfilename(
                parent=self.root,
                title="选择原来的批次配置",
                filetypes=(("批次配置", "*.json"), ("所有文件", "*.*")),
            )
            if not chosen:
                self.status_var.set("没有批次配置，已取消继续。")
                return
            config_path = Path(chosen)
        try:
            config = load_batch_config(config_path)
        except ValueError as exc:
            messagebox.showerror("无法读取批次配置", str(exc), parent=self.root)
            return
        if _normalized_path(config["record_path"]) != _normalized_path(record_path):
            messagebox.showerror(
                "配置与记录不一致",
                "所选批次配置对应的记录不是这份运行记录。程序不会改用窗口里的默认值重新导出。",
                parent=self.root,
            )
            return
        request = copy.deepcopy(config["request"])
        request["resume"] = True
        request["overwrite"] = bool(config["request"].get("overwrite", False))
        if retry:
            from .batch import load_batch_record

            record = load_batch_record(record_path)
            request["retry_task_ids"] = [
                task_id
                for task_id, task in dict(record.get("tasks") or {}).items()
                if isinstance(task, dict) and task.get("status") in {"failed", "cancelled", "blocked"}
            ]
        else:
            request["retry_task_ids"] = []
        self._saved_request = request
        self._arm_restored_batch(config, request)

    def _arm_restored_batch(self, config: dict[str, Any], request: dict[str, Any]) -> None:
        """Show the saved source list, then start that saved request."""
        self._apply_saved_request_to_widgets(config)
        self._set_busy(True)
        pending = [Path(item) for item in request["inputs"]]
        self._clear_listed_sources()
        self._paths = []
        self._path_keys = []
        self._source_errors.clear()
        self._loaded_plan = None
        self._plan_base_dir = None

        def step() -> None:
            if not self._ui_alive():
                return
            started = time.monotonic()
            inserted = 0
            while pending and inserted < _UI_ROW_BUDGET and (time.monotonic() - started) < _UI_FRAME_SECONDS:
                path = pending.pop(0)
                key = self._cached_path_key(path)
                self._paths.append(path)
                self._path_keys.append(key)
                self._remember_key(path, key)
                self._insert_listed_row({"path": str(path), "path_key": key})
                inserted += 1
            if pending:
                try:
                    self.status_var.set(f"正在恢复来源列表 {len(self._paths)}/{len(self._paths) + len(pending)}…")
                except tk.TclError:
                    return
                self._after(1, step)
                return
            self._set_busy(False)
            self._launch_batch(request, name=str(config.get("name") or ""))

        self.status_var.set("正在恢复已保存的批次配置和全部来源…")
        step()

    def _apply_saved_request_to_widgets(self, config: dict[str, Any]) -> None:
        request = config["request"]
        self.batch_name_var.set(str(config.get("name") or ""))
        self.layout_var.set("per_file")
        fmt = request.get("format", "opju")
        self.format_var.set(fmt if fmt in _FORMAT_SUFFIX else "opju")
        self.pdf_var.set(bool(request.get("pdf")))
        self.batch_overwrite_var.set(bool(request.get("overwrite")))
        self.keep_open_var.set(bool(request.get("keep_open")))
        self._apply_read_options(request.get("read_options") or {})
        plot = request.get("plot") or {}
        kind = str(plot.get("kind") or "auto")
        self.batch_kind_var.set(kind if kind in {"auto", "none", "line", "scatter", "line_symbol", "column"} else "auto")
        explicit = plot.get("x") not in (None, "auto") or plot.get("y") not in (None, "auto") or isinstance(plot.get("y"), list)
        self.plot_scope_var.set("明确列" if explicit else "自动（按每张表建议）")
        if explicit:
            self.explicit_x_var.set("" if plot.get("x") in (None, "auto") else str(plot.get("x")))
            y_value = plot.get("y")
            if isinstance(y_value, list):
                self.explicit_y_var.set(", ".join(str(item) for item in y_value))
            elif y_value in (None, "auto"):
                self.explicit_y_var.set("")
            else:
                self.explicit_y_var.set(str(y_value))
            errors = plot.get("y_error")
            if isinstance(errors, dict) and errors:
                self.explicit_error_var.set(str(next(iter(errors.values()))))
            else:
                self.explicit_error_var.set("")
        else:
            self.explicit_x_var.set("")
            self.explicit_y_var.set("")
            self.explicit_error_var.set("")
        self._sync_batch_controls()

    def export_data(self) -> None:
        if self._busy:
            return
        if self._per_file_mode():
            self._export_per_file()
            return
        fmt = self.format_var.get()
        if fmt not in _FORMAT_SUFFIX:
            messagebox.showerror("输出格式无效", "请选择 Origin 工程或 Excel 工作簿。", parent=self.root)
            return
        if not self._paths and self._loaded_plan is None:
            messagebox.showinfo("没有数据", "请先添加数据文件或加载导入计划。", parent=self.root)
            return
        if fmt == "opju":
            try:
                import spectra_to_origin as sto

                if sto.origin_process_running():
                    messagebox.showwarning("请先关闭 Origin", "检测到 Origin 正在运行。请先保存并关闭当前 Origin 工程，再重新创建文件；程序不会重置当前会话。", parent=self.root)
                    return
            except Exception as exc:
                messagebox.showerror("Origin 状态检查失败", str(exc), parent=self.root)
                return
        raw_output = filedialog.asksaveasfilename(
            parent=self.root,
            title="选择输出文件",
            defaultextension=_FORMAT_SUFFIX[fmt],
            initialfile=self._default_output_path(fmt).name,
            initialdir=str(self._default_output_path(fmt).parent),
            filetypes=(("Origin 工程", "*.opju"),) if fmt == "opju" else (("Excel 工作簿", "*.xlsx"),),
        )
        if not raw_output:
            return
        output = Path(raw_output).expanduser()
        if output.suffix.lower() != _FORMAT_SUFFIX[fmt]:
            output = output.with_suffix(_FORMAT_SUFFIX[fmt])
        if _normalized_path(output) in {_normalized_path(path) for path in self._paths}:
            messagebox.showerror("输出路径冲突", "输出文件不能与任何输入数据源相同。", parent=self.root)
            return
        overwrite = False
        if output.exists():
            if not messagebox.askyesno("确认覆盖", f"文件已存在：\n{output}\n\n确认覆盖该文件吗？", parent=self.root):
                return
            overwrite = True
        try:
            plan = self._effective_plan(output, fmt, overwrite)
        except Exception as exc:
            messagebox.showerror("无法创建导入计划", str(exc), parent=self.root)
            return
        self.status_var.set("正在后台准备和导出；窗口仍可响应。")
        try:
            self._start_import(plan)
        except Exception as exc:
            self._set_busy(False)
            messagebox.showerror("无法启动导出", str(exc), parent=self.root)
            self.status_var.set("无法启动后台导出。")

    def _start_import(self, plan: dict[str, Any]) -> None:
        if self._busy:
            return
        context = multiprocessing.get_context("spawn")
        recv_conn, send_conn = context.Pipe(duplex=False)
        base_dir = self._plan_base_dir or Path.cwd()
        process = context.Process(target=import_worker, args=(send_conn, plan, str(base_dir)))
        try:
            process.start()
        except Exception:
            recv_conn.close()
            send_conn.close()
            raise
        send_conn.close()
        self._import_recv = recv_conn
        self._import_process = process
        self._import_cancel = None
        self._import_kind = "plan"
        self._import_result = None
        self._import_exit_waits = 0
        self._import_eof = False
        self._set_busy(True)
        self._after(100, self._poll_import)

    def _poll_import(self) -> None:
        if not self._ui_alive():
            self._request_batch_cancel()
            self._release_import_handles()
            self._busy = False
            return
        process = self._import_process
        recv_conn = self._import_recv
        if process is None or recv_conn is None:
            return
        started = time.monotonic()
        taken = 0
        while taken < _IMPORT_MESSAGE_BUDGET and (time.monotonic() - started) < _IMPORT_FRAME_SECONDS:
            try:
                if not recv_conn.poll():
                    break
                message = recv_conn.recv()
            except (EOFError, OSError):
                self._import_eof = True
                break
            taken += 1
            if not isinstance(message, dict):
                continue
            if message.get("type") == "result":
                result = message.get("result")
                if isinstance(result, dict):
                    self._import_result = result
            elif message.get("type") == "progress":
                event = message.get("event")
                if isinstance(event, dict):
                    try:
                        self._apply_batch_progress(event)
                    except tk.TclError:
                        self._request_batch_cancel()
                        self._release_import_handles()
                        self._busy = False
                        return
                text = message.get("text")
                if text:
                    try:
                        self.status_var.set(str(text))
                    except tk.TclError:
                        self._request_batch_cancel()
                        self._release_import_handles()
                        self._busy = False
                        return
        try:
            alive = process.is_alive()
        except ValueError:
            alive = False
        if alive:
            self._import_exit_waits = 0
            self._after(120, self._poll_import)
            return
        if not self._import_eof:
            try:
                if recv_conn.poll():
                    self._after(1, self._poll_import)
                    return
            except (EOFError, OSError):
                self._import_eof = True
        if self._import_result is None and not self._import_eof and self._import_exit_waits < _IMPORT_EXIT_POLLS:
            self._import_exit_waits += 1
            self._after(30, self._poll_import)
            return
        self._complete_import()

    def _apply_batch_progress(self, event: dict[str, Any]) -> None:
        if event.get("type") == "batch":
            count = event.get("count")
            self.status_var.set(f"批量导出已开始，共 {count} 个来源。停止后续会在当前项结束后生效。")
            return
        if event.get("type") != "task":
            return
        tree = getattr(self, "task_tree", None)
        if tree is None:
            return
        task_id = str(event.get("task_id") or event.get("source") or "")
        source = str(event.get("source") or "")
        values = (source, _format_task_status(event), _format_task_detail(event))
        iid = self._task_iids.get(task_id)
        if iid and tree.exists(iid):
            tree.item(iid, values=values)
        else:
            self._task_seq = getattr(self, "_task_seq", 0) + 1
            iid = f"t{self._task_seq}"
            tree.insert("", tk.END, iid=iid, values=values)
            self._task_iids[task_id] = iid
        tree.see(iid)
        index = event.get("index")
        count = event.get("count")
        if isinstance(index, int) and isinstance(count, int):
            self.status_var.set(f"批量进度 {index + 1}/{count}：{Path(source).name} {_format_task_status(event)}")

    def _release_import_handles(self) -> None:
        """Close the parent pipe. Close a finished child; reap a live one later.

        Used when the window is already going away, including direct
        ``root.destroy()`` after the import poll has been cancelled. A worker
        that is still saving is not terminated and is not joined on the Tk
        thread. Batch close requests cancel before the parent pipe is closed
        so the current task can still finish and later tasks do not start.
        ``_complete_import`` still finishes a normal export.
        """
        self._request_batch_cancel()
        self._import_cancel = None
        self._import_kind = None
        process = self._import_process
        recv_conn = self._import_recv
        self._import_process = None
        self._import_recv = None
        if recv_conn is not None:
            try:
                recv_conn.close()
            except OSError:
                pass
        if process is None:
            return
        try:
            alive = bool(process.is_alive())
        except (ValueError, OSError):
            alive = False
        if alive:
            self._reap_import_process(process)
            return
        try:
            process.join(timeout=0)
        except (ValueError, OSError):
            pass
        try:
            process.close()
        except (ValueError, OSError):
            pass

    def _reap_import_process(self, process) -> None:
        """Join and close an export process after it exits, without touching Tk."""

        def reap() -> None:
            while True:
                try:
                    process.join(timeout=0.2)
                except (ValueError, OSError):
                    return
                try:
                    if not process.is_alive():
                        break
                except ValueError:
                    return
            try:
                process.join(timeout=0)
            except (ValueError, OSError):
                pass
            try:
                process.close()
            except (ValueError, OSError):
                pass

        threading.Thread(target=reap, name="origin-bridge-import-reap", daemon=True).start()

    def _complete_import(self) -> None:
        kind = getattr(self, "_import_kind", None)
        process = self._import_process
        recv_conn = self._import_recv
        self._import_process = None
        self._import_recv = None
        self._import_cancel = None
        self._import_kind = None
        exit_code = None
        if process is not None:
            try:
                if process.is_alive():
                    process.join(timeout=1)
            except ValueError:
                pass
            exit_code = getattr(process, "exitcode", None)
            try:
                process.close()
            except ValueError:
                pass
        if recv_conn is not None:
            try:
                recv_conn.close()
            except OSError:
                pass
        try:
            self._set_busy(False)
        except tk.TclError:
            self._busy = False
        if not self._ui_alive():
            return
        result = self._import_result if isinstance(self._import_result, dict) else {
            "ok": False,
            "error": f"后台进程未返回结果（退出码 {exit_code}）。",
        }
        if kind == "batch":
            self._present_batch_result(result)
            self._import_result = None
            return
        if result.get("ok"):
            output = result.get("output", {})
            warnings = result.get("warnings", [])
            path = output.get("path", "") if isinstance(output, dict) else str(output)
            size = output.get("size_bytes") if isinstance(output, dict) else None
            try:
                self.status_var.set(f"已创建：{path}" + (f"（{size:,} 字节）" if isinstance(size, int) else ""))
                details = f"文件已创建：\n{path}"
                if isinstance(size, int):
                    details += f"\n大小：{size:,} 字节"
                if warnings:
                    details += "\n\n提示：\n" + "\n".join(map(str, warnings))
                messagebox.showinfo("导入完成", details, parent=self.root)
            except tk.TclError:
                self._busy = False
        else:
            error = result.get("error", "导出失败")
            try:
                self.status_var.set(f"导出失败：{error}")
                messagebox.showerror("导入失败", str(error), parent=self.root)
            except tk.TclError:
                self._busy = False

    def save_plan(self) -> None:
        if self._busy:
            return
        if not self._paths and self._loaded_plan is None:
            messagebox.showinfo("没有数据", "请先添加数据文件或加载导入计划。", parent=self.root)
            return
        fmt = self.format_var.get()
        try:
            if self._loaded_plan and isinstance(self._loaded_plan.get("output"), dict):
                output = _resolve_plan_path(
                    self._loaded_plan["output"].get("path") or self._default_output_path(fmt),
                    self._plan_base_dir,
                )
            else:
                output = self._default_output_path(fmt)
            plan = self._effective_plan(output, fmt, False)
        except Exception as exc:
            messagebox.showerror("无法生成计划", str(exc), parent=self.root)
            return
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title="保存导入计划",
            defaultextension=".json",
            initialfile=f"{output.stem}_import_plan.json",
            filetypes=(("JSON 导入计划", "*.json"),),
        )
        if not path:
            return
        target = Path(path)
        if target.suffix.lower() != ".json":
            target = target.with_suffix(".json")
        try:
            _validate_plan_save_target(plan, target, self._plan_base_dir)
        except ValueError as exc:
            messagebox.showerror("不能保存计划", str(exc), parent=self.root)
            return
        if target.exists() and not messagebox.askyesno("确认覆盖", f"计划文件已存在：\n{target}\n\n确认覆盖吗？", parent=self.root):
            return
        try:
            saved_plan = copy.deepcopy(plan)
            for item in saved_plan["tables"]:
                item["source"]["path"] = str(_resolve_plan_path(item["source"]["path"], self._plan_base_dir))
            saved_plan["output"]["path"] = str(_resolve_plan_path(saved_plan["output"]["path"], self._plan_base_dir))
            target.write_text(json.dumps(saved_plan, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        except OSError as exc:
            messagebox.showerror("保存失败", str(exc), parent=self.root)
            return
        self.status_var.set(f"导入计划已保存：{target}")

    def load_plan(self) -> None:
        if self._busy:
            return
        path = filedialog.askopenfilename(
            parent=self.root,
            title="加载导入计划",
            filetypes=(("JSON 导入计划", "*.json"), ("所有文件", "*.*")),
        )
        if not path:
            return
        try:
            plan = json.loads(Path(path).read_text(encoding="utf-8-sig"))
            if not isinstance(plan, dict) or plan.get("schema_version") != PLAN_SCHEMA_VERSION:
                raise ValueError("计划 schema_version 必须为 1。")
            if not isinstance(plan.get("output"), dict) or not isinstance(plan.get("tables"), list):
                raise ValueError("计划必须包含 output 对象和 tables 数组。")
            if not plan["tables"]:
                raise ValueError("计划没有可导入的数据表。")
            for item in plan["tables"]:
                if not isinstance(item, dict) or not isinstance(item.get("source"), dict) or not item["source"].get("path"):
                    raise ValueError("计划中的每张表都必须包含 source.path。")
            self._plan_base_dir = Path(path).resolve().parent
            self.layout_var.set("unified")
            self._sync_batch_controls()
            self._loaded_plan = plan
            self._paths = list(dict.fromkeys(
                _resolve_plan_path(item["source"]["path"], self._plan_base_dir) for item in plan["tables"]
            ))
            self._path_keys = [self._cached_path_key(path) for path in self._paths]
            self._plots.clear()
            self._column_labels.clear()
            self._fingerprints.clear()
            self._source_errors.clear()
            for item in plan["tables"]:
                key = _plan_table_key(item, self._plan_base_dir)
                if isinstance(item.get("plot"), dict):
                    self._plots[key] = copy.deepcopy(item["plot"])
                if isinstance(item.get("column_labels"), dict):
                    self._column_labels[key] = copy.deepcopy(item["column_labels"])
                source = item["source"]
                self._fingerprints[key] = (
                    source.get("sha256"),
                    options_fingerprint(source.get("options") or {}),
                )
            self._clear_listed_sources()
            output = plan["output"]
            fmt = output.get("format", "opju")
            self.format_var.set(fmt if fmt in _FORMAT_SUFFIX else "opju")
            self._format_changed()
            self.keep_open_var.set(bool(output.get("keep_open", False)))
            options = plan["tables"][0]["source"].get("options") or {}
            self._apply_read_options(options)
            self._inspect_plan_sources(plan)
        except Exception as exc:
            messagebox.showerror("无法加载计划", str(exc), parent=self.root)

    def _apply_read_options(self, options: dict[str, Any]) -> None:
        header = options.get("header", "auto")
        self.header_var.set("自动" if header == "auto" else "有标题行" if header is True else "无标题行")
        self.skip_rows_var.set(str(options.get("skip_rows", 0)))
        delimiter = options.get("delimiter", "auto")
        label = {"auto": "自动识别", ",": "逗号 ,", "\t": "制表符 \\t", ";": "分号 ;", "whitespace": "空白分隔"}.get(delimiter, "自动识别")
        self.delimiter_var.set(label)

    def open_spectra_gui(self) -> None:
        if self._busy:
            return
        try:
            if getattr(sys, "frozen", False):
                command = [sys.executable, "--spectra-gui"]
            else:
                script = Path(__file__).resolve().parent.parent / "spectra_to_origin.py"
                command = [sys.executable, str(script), "--spectra-gui"]
            subprocess.Popen(command, cwd=str(Path.cwd()))
        except Exception as exc:
            messagebox.showerror("无法打开批量谱线模式", str(exc), parent=self.root)

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        for widget in self._interactive:
            state = "disabled" if busy else self._widget_states.get(widget, "normal")
            try:
                widget.configure(state=state)
            except tk.TclError:
                continue
        self.y_list.configure(state=tk.DISABLED if busy else tk.NORMAL)
        try:
            self.table_tree.state(["disabled"] if busy else ["!disabled"])
        except tk.TclError:
            pass
        self._sync_stop_button()

    def _present_batch_result(self, result: dict[str, Any]) -> None:
        tasks = result.get("tasks") if isinstance(result.get("tasks"), list) else []
        tree = getattr(self, "task_tree", None)
        if tree is not None and tasks and not tree.get_children():
            for task in tasks:
                if not isinstance(task, dict):
                    continue
                output = task.get("output")
                event = dict(task)
                event["type"] = "task"
                if isinstance(output, dict):
                    event["output"] = output.get("path")
                self._apply_batch_progress(event)
        counts = result.get("counts") if isinstance(result.get("counts"), dict) else {}
        error = result.get("error")
        if error and not counts:
            text = error if isinstance(error, str) else str(error)
            try:
                self.status_var.set(f"批量导出失败：{text}")
                if not os.environ.get("ORI_GUI_BATCH_DRIVE"):
                    messagebox.showerror("批量导出失败", text, parent=self.root)
            except tk.TclError:
                self._busy = False
            return
        text = "批量导出已结束。" + _format_batch_counts(counts)
        if result.get("cancelled") or int(counts.get("cancelled") or 0):
            text += "。已成功的文件保留，尚未开始的任务已取消。"
        try:
            self.status_var.set(text)
            if not os.environ.get("ORI_GUI_BATCH_DRIVE"):
                messagebox.showinfo("批量导出结束", text + "\n任务列表里有每个文件的状态；PDF 失败会写出具体的图和原因。", parent=self.root)
        except tk.TclError:
            self._busy = False

    def batch_source_snapshot(self) -> dict[str, Any]:
        """Paths plus light summaries. Failed sources stay in ``errors``.

        Per-file export may include paths that have not been inspected. This
        snapshot does not require every source to have succeeded.
        """
        store = getattr(self, "_summaries", None)
        return {
            "paths": [str(path) for path in self._paths],
            "tables": [table for table in self._tables if not table.get("error")],
            "errors": dict(getattr(self, "_source_errors", {}) or {}),
            "inspect_calls": getattr(store, "inspect_calls", 0),
            "cache_hits": getattr(store, "cache_hits", 0),
        }

    def _on_close(self) -> None:
        if self._busy and getattr(self, "_import_kind", None) == "batch":
            messagebox.showwarning(
                "批量仍在运行",
                "当前这项会继续保存。“停止后续”只取消还没开始的任务，不能打断卡在启动证明之前的 Origin。\n\n"
                "请先点“停止后续”，等任务列表停止后再关闭。关闭窗口不会强制结束正在写的任务。",
                parent=self.root,
            )
            return
        if self._busy:
            messagebox.showwarning("任务进行中", "数据读取或导出仍在后台运行，请等待完成后再关闭窗口。", parent=self.root)
            return
        self.root.destroy()

    def _schedule_drive(self) -> None:
        spec = os.environ.get("ORI_GUI_BATCH_DRIVE")
        if not spec or multiprocessing.current_process().name != "MainProcess":
            return

        def begin() -> None:
            if not self._ui_alive():
                return
            from .gui_drive import run_drive

            run_drive(self, spec)

        self._after(300, begin)


__all__ = ["GeneralDataApp"]
