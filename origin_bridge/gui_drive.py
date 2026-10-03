"""Hidden startup path for exercising the packaged general-data window.

Set ``ORI_GUI_BATCH_DRIVE`` to a JSON file. Nothing in the normal window
calls this, and the variable is removed before a batch child is spawned so
the frozen executable does not open another window. The drive uses the same
batch launch, cancel, and resume methods as the buttons.
"""

from __future__ import annotations

import ctypes
import json
import os
import time
from pathlib import Path
from typing import Any, Callable

from .batch_config import config_path_for, load_batch_config


def run_drive(app: Any, spec_path: str) -> None:
    report: dict[str, Any] = {
        "ok": False,
        "parent_pid": os.getpid(),
        "child_pids": [],
        "window_count_max": 0,
        "inspect_calls_before_export": None,
        "heartbeat_max_gap_s": None,
        "screenshots": {},
        "phases": [],
    }
    spec = json.loads(Path(spec_path).read_text(encoding="utf-8"))
    try:
        _run(app, spec, report)
        report["ok"] = True
    except Exception as exc:
        report["error"] = f"{exc.__class__.__name__}: {exc}"
    finally:
        target = Path(spec["report"])
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        try:
            if app.root.winfo_exists():
                app.root.destroy()
        except Exception:
            pass


def _run(app: Any, spec: dict[str, Any], report: dict[str, Any]) -> None:
    app.root.deiconify()
    app.root.geometry("1100x860+40+40")
    app.root.lift()
    try:
        app.root.attributes("-topmost", True)
    except Exception:
        pass
    app.root.update()
    shots = spec.get("screenshots") or {}
    app.layout_var.set("per_file")
    app.format_var.set("xlsx")
    app.batch_kind_var.set("none")
    app.plot_scope_var.set("自动（按每张表建议）")
    app.batch_name_var.set(str(spec.get("name") or "frozen-batch"))
    app.pdf_var.set(False)
    app.batch_overwrite_var.set(False)
    app._add_paths(list(spec["inputs"]))
    gaps = _pump(app, lambda: not app._busy, 180, report)
    report["inspect_calls_before_export"] = int(app._summaries.inspect_calls)
    if app._summaries.inspect_calls != 0:
        raise RuntimeError(f"per-file discovery inspected {app._summaries.inspect_calls} sources")
    if len(app._paths) != len(spec["inputs"]):
        raise RuntimeError(f"discovered {len(app._paths)} sources, expected {len(spec['inputs'])}")
    _note_windows(app, report)
    from .batch import build_batch_request

    output_dir = Path(spec["output_dir"])
    request = build_batch_request(
        app._paths,
        output_dir,
        layout="per_file",
        format="xlsx",
        read_options={"header": "auto", "skip_rows": 0, "delimiter": "auto", "encoding": "auto", "missing_values": [""], "formula_policy": "cached"},
        plot={"kind": "none", "x": "auto", "y": "auto", "y_error": "auto", "title": "auto", "x_label": "auto", "y_label": "auto"},
        overwrite=False,
        pdf=False,
        keep_open=False,
        resume=False,
    )
    app._saved_request = request
    app._launch_batch(request, name=str(spec.get("name") or ""))
    _pump(app, lambda: bool(app.task_tree.get_children()) or not app._busy, 180, report)
    _note_child(app, report)
    _note_windows(app, report)
    _shot(app, shots.get("started"), report, "started")
    app.stop_later_tasks()
    gaps.extend(_pump(app, lambda: not app._busy, 180, report))
    _shot(app, shots.get("cancelled"), report, "cancelled")
    report["phases"].append({"name": "cancel", "status": app.status_var.get()})
    _resume(app, request["record_path"], retry=False)
    gaps.extend(_pump(app, lambda: not app._busy, 180, report))
    _note_child(app, report)
    _shot(app, shots.get("resumed"), report, "resumed")
    report["phases"].append({"name": "resume", "status": app.status_var.get()})
    record = _load_record(request["record_path"])
    failed = [
        task for task in dict(record.get("tasks") or {}).values()
        if isinstance(task, dict) and task.get("status") in {"failed", "cancelled", "blocked"}
    ]
    if failed and spec.get("repair_path"):
        Path(spec["repair_path"]).write_text("x,y\n1,2\n3,4\n", encoding="utf-8")
        _resume(app, request["record_path"], retry=True)
        gaps.extend(_pump(app, lambda: not app._busy, 180, report))
        _note_child(app, report)
        _shot(app, shots.get("retried"), report, "retried")
        report["phases"].append({"name": "retry", "status": app.status_var.get()})
    report["heartbeat_max_gap_s"] = max(gaps) if gaps else None
    report["final_status"] = app.status_var.get()
    _note_windows(app, report)


def _resume(app: Any, record_path: str, *, retry: bool) -> None:
    config = load_batch_config(config_path_for(record_path))
    request = json.loads(json.dumps(config["request"]))
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
    app._arm_restored_batch(config, request)


def _load_record(record_path: str) -> dict[str, Any]:
    path = Path(record_path)
    if not path.is_file():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    return data if isinstance(data, dict) else {}


def _pump(app: Any, done: Callable[[], bool], timeout: float, report: dict[str, Any]) -> list[float]:
    gaps: list[float] = []
    previous = time.monotonic()
    deadline = previous + timeout
    ticks = 0
    while time.monotonic() < deadline:
        now = time.monotonic()
        gaps.append(now - previous)
        previous = now
        app.root.update()
        ticks += 1
        if ticks % 20 == 0:
            _note_child(app, report)
            _note_windows(app, report)
        if done():
            app.root.update()
            _note_child(app, report)
            _note_windows(app, report)
            return gaps
        time.sleep(0.01)
    raise TimeoutError("frozen batch window did not reach the expected state")


def _note_child(app: Any, report: dict[str, Any]) -> None:
    process = getattr(app, "_import_process", None)
    if process is None:
        return
    pid = getattr(process, "pid", None)
    if isinstance(pid, int) and pid not in report["child_pids"]:
        report["child_pids"].append(pid)


def _note_windows(app: Any, report: dict[str, Any]) -> None:
    count = _visible_title_count("数据 → Origin 工程")
    report["window_count_max"] = max(int(report.get("window_count_max") or 0), count)


def _shot(app: Any, path: str | None, report: dict[str, Any], name: str) -> None:
    if not path:
        return
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    capture_tk(app.root, target)
    report["screenshots"][name] = str(target)


def capture_tk(widget: Any, path: Path) -> None:
    """Save a BMP of the top-level window without extra imaging packages."""
    widget.update_idletasks()
    user32 = ctypes.windll.user32
    hwnd = int(widget.winfo_id())
    root = user32.GetAncestor(hwnd, 2) or hwnd
    user32.SetForegroundWindow(root)
    rect = ctypes.wintypes.RECT()
    user32.GetWindowRect(root, ctypes.byref(rect))
    width = max(1, int(rect.right - rect.left))
    height = max(1, int(rect.bottom - rect.top))
    hwnd_dc = user32.GetWindowDC(root)
    gdi32 = ctypes.windll.gdi32
    mem_dc = gdi32.CreateCompatibleDC(hwnd_dc)
    bitmap = gdi32.CreateCompatibleBitmap(hwnd_dc, width, height)
    gdi32.SelectObject(mem_dc, bitmap)
    user32.PrintWindow(root, mem_dc, 2)
    stride = ((width * 3 + 3) // 4) * 4
    header = _bitmap_info(width, height)
    pixels = ctypes.create_string_buffer(stride * height)
    gdi32.GetDIBits(mem_dc, bitmap, 0, height, pixels, ctypes.byref(header), 0)
    _write_bmp(path, width, height, stride, pixels.raw)
    gdi32.DeleteObject(bitmap)
    gdi32.DeleteDC(mem_dc)
    user32.ReleaseDC(root, hwnd_dc)


class _BitmapInfoHeader(ctypes.Structure):
    _pack_ = 1
    _fields_ = (
        ("biSize", ctypes.c_uint32),
        ("biWidth", ctypes.c_int32),
        ("biHeight", ctypes.c_int32),
        ("biPlanes", ctypes.c_uint16),
        ("biBitCount", ctypes.c_uint16),
        ("biCompression", ctypes.c_uint32),
        ("biSizeImage", ctypes.c_uint32),
        ("biXPelsPerMeter", ctypes.c_int32),
        ("biYPelsPerMeter", ctypes.c_int32),
        ("biClrUsed", ctypes.c_uint32),
        ("biClrImportant", ctypes.c_uint32),
    )


def _bitmap_info(width: int, height: int) -> _BitmapInfoHeader:
    header = _BitmapInfoHeader()
    header.biSize = ctypes.sizeof(_BitmapInfoHeader)
    header.biWidth = width
    header.biHeight = height
    header.biPlanes = 1
    header.biBitCount = 24
    header.biCompression = 0
    return header


def _write_bmp(path: Path, width: int, height: int, stride: int, pixels: bytes) -> None:
    pixel_size = stride * height
    file_size = 14 + 40 + pixel_size
    header = b"BM" + file_size.to_bytes(4, "little") + (0).to_bytes(4, "little") + (54).to_bytes(4, "little")
    info = _bitmap_info(width, height)
    path.write_bytes(header + bytes(info) + pixels[:pixel_size])


def _visible_title_count(title: str) -> int:
    user32 = ctypes.windll.user32
    found: list[int] = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def visit(hwnd, _lparam):
        if not user32.IsWindowVisible(hwnd):
            return True
        length = user32.GetWindowTextLengthW(hwnd)
        if length <= 0:
            return True
        buffer = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buffer, length + 1)
        if title in buffer.value:
            found.append(int(hwnd))
        return True

    user32.EnumWindows(visit, 0)
    return len(found)
