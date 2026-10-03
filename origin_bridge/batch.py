"""Bounded batch export for one OPJU/XLSX per input file, plus the legacy unified project.

Requests and results are JSON and pickle safe. Full table bodies stay in a small
prefetch window. A success record is reused only when the source bytes, read
settings, plot settings, and output bytes still match. File existence, size, and
mtime are not enough.

GUI code should call :func:`build_batch_request` and :func:`execute_batch`.
The spawn entry point is :func:`origin_bridge.worker.batch_worker`.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import tempfile
from collections import OrderedDict
from dataclasses import replace
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Mapping

from .models import DataImportError, DataTable, PlannedTable, PreparedImport
from .session import OriginSession, OriginSessionLost

BATCH_SCHEMA_VERSION = 1
RECORD_KIND = "ori-batch-record"
_FORMAT_SUFFIX = {"opju": ".opju", "xlsx": ".xlsx"}
_PLOT_KINDS = frozenset({"auto", "none", "line", "scatter", "line_symbol", "column"})
_INVALID_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_GENERATED_NAME = re.compile(
    r"__[0-9a-f]{12}(?:__g\d{2}_.*)?\.(?:opju|xlsx|pdf)$",
    re.IGNORECASE,
)
_DEFAULT_READ_OPTIONS: dict[str, Any] = {
    "header": "auto",
    "skip_rows": 0,
    "delimiter": "auto",
    "encoding": "auto",
    "missing_values": [""],
    "formula_policy": "cached",
}
_DEFAULT_PLOT: dict[str, Any] = {
    "kind": "auto",
    "x": "auto",
    "y": "auto",
    "y_error": "auto",
    "title": "auto",
    "x_label": "auto",
    "y_label": "auto",
}
_TERMINAL = frozenset({"succeeded", "failed", "skipped", "cancelled", "blocked"})


def build_batch_request(
    inputs: list[str | Path],
    output_dir: str | Path,
    *,
    layout: str = "per_file",
    format: str = "opju",
    read_options: Mapping[str, Any] | None = None,
    plot: Mapping[str, Any] | None = None,
    task_overrides: list[Mapping[str, Any]] | None = None,
    overwrite: bool = False,
    pdf: bool = False,
    keep_open: bool = False,
    recycle_every: int = 0,
    prefetch: int = 1,
    max_resident_tasks: int = 2,
    origin_timeout_s: float = 300,
    origin_retries: int = 1,
    output_name: str | None = None,
    record_path: str | Path | None = None,
    resume: bool = False,
    retry_task_ids: list[str] | None = None,
) -> dict[str, Any]:
    """Build a JSON-safe batch request. This does not read data or start Origin."""
    if layout not in {"per_file", "unified"}:
        raise DataImportError("layout must be 'per_file' or 'unified'")
    if format not in _FORMAT_SUFFIX:
        raise DataImportError("format must be 'opju' or 'xlsx'")
    if pdf and format != "opju":
        raise DataImportError("pdf export requires format 'opju'")
    if keep_open and format != "opju":
        raise DataImportError("keep_open is only valid for opju output")
    if not isinstance(overwrite, bool) or not isinstance(pdf, bool) or not isinstance(keep_open, bool):
        raise DataImportError("overwrite, pdf, and keep_open must be booleans")
    if not isinstance(resume, bool):
        raise DataImportError("resume must be a boolean")
    for name, value, lower in (
        ("recycle_every", recycle_every, 0),
        ("prefetch", prefetch, 0),
        ("origin_retries", origin_retries, 0),
    ):
        if isinstance(value, bool) or not isinstance(value, int) or value < lower:
            raise DataImportError(f"{name} must be an integer >= {lower}")
    if isinstance(max_resident_tasks, bool) or not isinstance(max_resident_tasks, int) or max_resident_tasks < 1:
        raise DataImportError("max_resident_tasks must be an integer >= 1")
    if isinstance(origin_timeout_s, bool) or not isinstance(origin_timeout_s, (int, float)) or origin_timeout_s <= 0:
        raise DataImportError("origin_timeout_s must be a positive number")
    if not inputs:
        raise DataImportError("At least one input path is required")
    directory = _absolute(output_dir)
    if directory.exists() and not directory.is_dir():
        raise DataImportError(f"output_dir is not a directory: {directory}")
    record = _absolute(record_path) if record_path else directory / "batch-record.json"
    request = {
        "schema_version": BATCH_SCHEMA_VERSION,
        "inputs": [str(_absolute(item)) for item in inputs],
        "output_dir": str(directory),
        "layout": layout,
        "format": format,
        "read_options": _merge_options(_DEFAULT_READ_OPTIONS, read_options),
        "plot": _merge_plot(_DEFAULT_PLOT, plot),
        "task_overrides": [_public_override(item) for item in (task_overrides or [])],
        "overwrite": overwrite,
        "pdf": pdf,
        "keep_open": keep_open,
        "recycle_every": recycle_every,
        "prefetch": prefetch,
        "max_resident_tasks": max_resident_tasks,
        "origin_timeout_s": float(origin_timeout_s),
        "origin_retries": origin_retries,
        "output_name": output_name or f"batch{_FORMAT_SUFFIX[format]}",
        "record_path": str(record),
        "resume": resume,
        "retry_task_ids": [str(item) for item in (retry_task_ids or [])],
    }
    _validate_plot_request(request["plot"], "plot")
    for index, item in enumerate(request["task_overrides"]):
        if "plot" in item:
            _validate_plot_request(item["plot"], f"task_overrides[{index}].plot", partial=True)
    _reject_unsupported_modes(request)
    json.dumps(request, ensure_ascii=False)
    return request


def execute_batch(
    request: Mapping[str, Any],
    *,
    progress: Callable[[dict[str, Any]], None] | None = None,
    cancel_event: Any = None,
    session_factory: Callable[[], Any] | None = None,
    resume: bool | None = None,
) -> dict[str, Any]:
    """Run a batch request. Failures stay on their own task. Cancel stops later tasks."""
    from .exporter import origin_export_lock

    job = _checked_request(request)
    if resume is not None:
        job["resume"] = bool(resume)
    _reject_unsupported_modes(job)
    with origin_export_lock():
        if job["layout"] == "unified":
            return _execute_unified(job, progress=progress, cancel_event=cancel_event)
        return _execute_per_file(
            job,
            progress=progress,
            cancel_event=cancel_event,
            session_factory=session_factory,
        )


def load_batch_record(path: str | Path) -> dict[str, Any]:
    """Load a record. A torn primary file falls back to the last complete backup."""
    record_path = Path(path)
    primary = _read_record_file(record_path)
    backup = _read_record_file(Path(str(record_path) + ".bak"))
    if primary is not None and not primary.get("_corrupt"):
        return primary
    if backup is not None and not backup.get("_corrupt"):
        backup["_recovered_from_backup"] = True
        return backup
    empty = _empty_record()
    if (primary and primary.get("_corrupt")) or (backup and backup.get("_corrupt")):
        empty["_corrupt"] = True
        empty["_corrupt_error"] = (primary or backup or {}).get("_corrupt_error", "record is not valid JSON")
    return empty


def save_batch_record(path: str | Path, record: Mapping[str, Any]) -> None:
    """Atomically replace a record and keep the previous complete file as ``.bak``."""
    record_path = Path(path)
    payload = {key: value for key, value in dict(record).items() if not str(key).startswith("_")}
    payload["schema_version"] = BATCH_SCHEMA_VERSION
    payload["kind"] = RECORD_KIND
    payload["updated_at"] = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    _atomic_json(record_path, payload)


def read_batch_log(path: str | Path) -> list[dict[str, Any]]:
    """Read a batch log. A torn final line is returned as a corrupt-line event."""
    log_path = Path(path)
    if not log_path.is_file():
        return []
    events: list[dict[str, Any]] = []
    text = log_path.read_text(encoding="utf-8", errors="replace")
    for line in text.splitlines():
        if not line.strip():
            continue
        try:
            item = json.loads(line)
        except json.JSONDecodeError:
            events.append({"type": "corrupt_log_line", "raw": line[:200]})
            continue
        if isinstance(item, dict):
            events.append(item)
        else:
            events.append({"type": "corrupt_log_line", "raw": line[:200]})
    return events


def retry_batch_tasks(
    record_path: str | Path,
    *,
    task_ids: list[str] | None = None,
    statuses: tuple[str, ...] = ("failed", "cancelled", "blocked"),
    force_task_ids: list[str] | None = None,
) -> dict[str, Any]:
    """Drop selected record entries so the next resumed run executes those tasks again.

    Output files are not deleted. A dropped success is therefore not skipped, and
    with ``overwrite`` false the new run blocks instead of replacing that file.
    """
    record = load_batch_record(record_path)
    selected = set(task_ids or [])
    forced = set(force_task_ids or [])
    kept: dict[str, Any] = {}
    for task_id, task in dict(record.get("tasks") or {}).items():
        source = str(task.get("source") or "")
        named = task_id in selected or source in selected or task_id in forced or source in forced
        if task_id in forced or source in forced:
            continue
        if task_ids is None and task.get("status") in statuses:
            continue
        if task_ids is not None and named and task.get("status") in statuses:
            continue
        kept[task_id] = task
    record["tasks"] = kept
    save_batch_record(record_path, record)
    return load_batch_record(record_path)


def log_path_for(record_path: str | Path) -> Path:
    path = Path(record_path)
    return path.with_suffix(".log.jsonl")


class _Resident:
    """Holds parsed tables for the current task plus a short prefetch, then drops them."""

    def __init__(self, limit: int) -> None:
        self.limit = max(1, limit)
        self.items: OrderedDict[str, Any] = OrderedDict()
        self.held = 0
        self.max_held = 0

    def note(self) -> None:
        self.max_held = max(self.max_held, self.held + len(self.items))

    def put(self, key: str, value: Any) -> bool:
        if key in self.items:
            self.items.move_to_end(key)
            self.items[key] = value
            self.note()
            return True
        while self.items and self.held + len(self.items) >= self.limit:
            self.items.popitem(last=False)
        if self.held + len(self.items) >= self.limit:
            self.note()
            return False
        self.items[key] = value
        self.note()
        return True

    def take(self, key: str) -> Any:
        value = self.items.pop(key, None)
        if value is not None:
            self.held += 1
            self.note()
        return value

    def hold_new(self) -> None:
        self.held += 1
        self.note()

    def release(self) -> None:
        if self.held:
            self.held -= 1


def _execute_per_file(
    job: dict[str, Any],
    *,
    progress: Callable[[dict[str, Any]], None] | None,
    cancel_event: Any,
    session_factory: Callable[[], Any] | None,
) -> dict[str, Any]:
    output_dir = Path(job["output_dir"])
    output_dir.mkdir(parents=True, exist_ok=True)
    record_path = Path(job["record_path"])
    tasks, excluded = _discover_tasks(job)
    record, record_warning = _open_record(record_path, resume=job["resume"], retry_ids=job["retry_task_ids"])
    resident = _Resident(int(job["max_resident_tasks"]))
    warnings = list(excluded)
    if record_warning:
        warnings.append(record_warning)
    state = {
        "counts": {name: 0 for name in ("succeeded", "skipped", "failed", "cancelled", "blocked", "pdf_failed")},
        "origin_starts": 0,
        "origin_stops": 0,
        "owned_pids": [],
        "owned_processes": [],
        "recycled": 0,
        "cancelled": False,
    }
    results: list[dict[str, Any]] = []
    session: Any = None
    since_start = 0
    window = max(1, min(int(job["prefetch"]) + 1, int(job["max_resident_tasks"])))
    _emit(record_path, progress, warnings, {"type": "batch", "status": "started", "count": len(tasks)})

    def prefetch_through(index: int) -> None:
        look = index
        while look < len(tasks) and resident.held + len(resident.items) < window:
            ahead = tasks[look]
            look += 1
            if ahead["task_id"] in resident.items:
                continue
            if ahead["task_id"] != tasks[index]["task_id"] and _task_decision(record, ahead, job)["kind"] != "run":
                continue
            try:
                loaded = _load_tables(Path(ahead["source"]), ahead["read_options"])
            except Exception as exc:
                resident.put(ahead["task_id"], ("err", exc, "", ahead["read_options"]))
                continue
            if not resident.put(ahead["task_id"], ("ok", loaded[0], loaded[1], ahead["read_options"])):
                del loaded
                break

    def close_session() -> None:
        nonlocal session
        if session is None:
            return
        before = int(getattr(session, "stops", 0))
        try:
            session.close()
        finally:
            state["origin_stops"] += int(getattr(session, "stops", 0)) - before
            session = None

    try:
        for index, task in enumerate(tasks):
            if _cancelled(cancel_event):
                state["cancelled"] = True
                _finish_task(record, record_path, results, state, task, index, len(tasks), progress, warnings,
                             status="cancelled", error=_error("Cancelled", "cancelled before this task started"))
                continue
            decision = _task_decision(record, task, job)
            if decision["kind"] == "skip":
                _finish_task(
                    record, record_path, results, state, task, index, len(tasks), progress, warnings,
                    status="skipped",
                    output=_stored_output(record, task["task_id"]),
                    pdfs=_stored_pdfs(record, task["task_id"]),
                    pdf_status=_stored_pdf_status(record, task["task_id"], job["pdf"]),
                )
                continue
            if decision["kind"] == "block":
                _report_preserved_block(
                    record_path, results, state, task, index, len(tasks), progress, warnings,
                    decision["reason"],
                )
                continue
            prefetch_through(index)
            if decision["kind"] == "recover":
                outcome = _recover_one(
                    job, task, decision, state, session, session_factory, since_start,
                )
            else:
                outcome = _run_one(
                    job,
                    task,
                    resident,
                    state,
                    session,
                    session_factory,
                    since_start,
                )
            session = outcome["session"]
            since_start = outcome["since_start"]
            if outcome.get("preserve_record"):
                _report_preserved_block(
                    record_path, results, state, task, index, len(tasks), progress, warnings,
                    (outcome.get("error") or {}).get("message") or "PDF recovery was refused",
                    pdf_errors=outcome.get("pdf_errors") or [],
                )
            else:
                _finish_task(
                    record, record_path, results, state, task, index, len(tasks), progress, warnings,
                    status=outcome["status"],
                    error=outcome.get("error"),
                    output=outcome.get("output"),
                    pdfs=outcome.get("pdfs") or [],
                    pdf_status=outcome.get("pdf_status", "not_applicable"),
                    source_sha256=outcome.get("source_sha256"),
                    warnings_for_task=outcome.get("warnings") or [],
                    resolved_plot=outcome.get("resolved_plot"),
                    pdf_errors=outcome.get("pdf_errors") or [],
                    graph_identities=outcome.get("graph_identities") or [],
                )
            if outcome.get("recycle"):
                close_session()
                since_start = 0
                state["recycled"] += 1
        if _cancelled(cancel_event):
            state["cancelled"] = True
    finally:
        if not job["keep_open"]:
            close_session()
        elif session is not None and not session.healthy():
            close_session()
        resident.items.clear()
        resident.held = 0

    return _result(job, results, state, warnings, resident.max_held)


def _run_one(job, task, resident: _Resident, state, session, session_factory, since_start) -> dict[str, Any]:
    cached = resident.take(task["task_id"])
    held_here = cached is not None
    tables: list[DataTable] | None = None
    try:
        if cached is None:
            resident.hold_new()
            held_here = True
            try:
                tables, source_hash = _load_tables(Path(task["source"]), task["read_options"])
            except Exception as exc:
                return _failed(exc, session, since_start)
        elif cached[0] == "err":
            return _failed(cached[1], session, since_start)
        else:
            _kind, tables, source_hash, cached_options = cached
            current_hash = _sha256_file(Path(task["source"]))
            if current_hash != source_hash or cached_options != task["read_options"]:
                del tables
                tables, source_hash = _load_tables(Path(task["source"]), task["read_options"])
        assert tables is not None
        output = Path(task["output"])
        try:
            prepared = _prepare_tables(tables, output, job["format"], job["overwrite"], task["plot"])
        except Exception as exc:
            return _failed(exc, session, since_start, source_hash)
        resolved = [_plot_public(item) for item in prepared.tables]
        if job["format"] == "xlsx":
            try:
                from .exporter import execute_import

                receipt = execute_import(prepared)
            except FileExistsError as exc:
                return _blocked(exc, session, since_start, source_hash)
            except Exception as exc:
                return _failed(exc, session, since_start, source_hash)
            return {
                "status": "succeeded",
                "session": session,
                "since_start": since_start,
                "output": receipt["output"],
                "pdfs": [],
                "pdf_status": "not_applicable",
                "source_sha256": source_hash,
                "warnings": list(receipt.get("warnings") or []),
                "resolved_plot": resolved,
            }
        graph_count = sum(item.plot.kind != "none" for item in prepared.tables)
        attempts = 0
        while True:
            if session is None or not session.healthy():
                if session is not None:
                    before_stops = int(getattr(session, "stops", 0))
                    session.close()
                    state["origin_stops"] += int(getattr(session, "stops", 0)) - before_stops
                session = _make_session(job, session_factory)
                before_starts = int(getattr(session, "starts", 0))
                try:
                    session.start()
                except Exception as exc:
                    state["origin_starts"] += int(getattr(session, "starts", 0)) - before_starts
                    return _failed(exc, session, since_start, source_hash)
                state["origin_starts"] += int(getattr(session, "starts", 0)) - before_starts
                _remember_owned(state, session)
                since_start = 0
            from .exporter import _all_warnings, _install_staged_file, _validate_prepared

            stage_dir = Path(tempfile.mkdtemp(prefix=f".{output.stem}_batch_", dir=str(output.parent)))
            try:
                try:
                    _tables, _output, converted, conversion_warnings, conversion_notes = _validate_prepared(prepared)
                except FileExistsError as exc:
                    return _blocked(exc, session, since_start, source_hash, resolved)
                except Exception as exc:
                    return _failed(exc, session, since_start, source_hash, resolved)
                try:
                    written = session.write_project(
                        stage_dir,
                        prepared.tables,
                        prepared,
                        converted,
                        pdf=bool(job["pdf"]) and graph_count > 0,
                        conversion_notes=conversion_notes,
                    )
                except FileExistsError as exc:
                    return _blocked(exc, session, since_start, source_hash, resolved)
                except Exception as exc:
                    if _is_session_loss(exc) and attempts < int(job["origin_retries"]):
                        attempts += 1
                        before_stops = int(getattr(session, "stops", 0))
                        session.close()
                        state["origin_stops"] += int(getattr(session, "stops", 0)) - before_stops
                        session = None
                        continue
                    return _failed(exc, session, since_start, source_hash, resolved)
                staged = Path(written["opju"])
                try:
                    _install_staged_file(staged, output, overwrite=bool(job["overwrite"]))
                except FileExistsError as exc:
                    return _blocked(exc, session, since_start, source_hash, resolved)
                pdfs, pdf_status, pdf_errors = _install_pdfs(
                    output, written.get("pdfs") or [], job, graph_count,
                )
                identities = _graph_identities(written.get("pdfs") or [])
                since_start += 1
                recycle = bool(job["recycle_every"]) and since_start >= int(job["recycle_every"])
                if written.get("session_lost"):
                    session_lost_obj = session
                    session = None
                    if session_lost_obj is not None:
                        before_stops = int(getattr(session_lost_obj, "stops", 0))
                        session_lost_obj.close()
                        state["origin_stops"] += int(getattr(session_lost_obj, "stops", 0)) - before_stops
                return {
                    "status": "succeeded",
                    "session": session,
                    "since_start": 0 if recycle or written.get("session_lost") else since_start,
                    "recycle": recycle and not written.get("session_lost"),
                    "output": {
                        "path": str(output),
                        "format": "opju",
                        "size_bytes": output.stat().st_size,
                        "sha256": _sha256_file(output),
                    },
                    "pdfs": pdfs,
                    "pdf_status": pdf_status,
                    "pdf_errors": pdf_errors,
                    "graph_identities": identities,
                    "source_sha256": source_hash,
                    "warnings": _all_warnings(prepared, prepared.tables, conversion_warnings),
                    "resolved_plot": resolved,
                }
            finally:
                shutil.rmtree(stage_dir, ignore_errors=True)
    finally:
        tables = None
        cached = None
        if held_here:
            resident.release()


def _install_pdfs(opju: Path, raw_pdfs: list[dict[str, Any]], job: Mapping[str, Any], graph_count: int):
    from .exporter import _install_staged_file
    from .session import validate_pdf_file

    if not job["pdf"] or graph_count == 0:
        return [], "not_applicable", []
    installed = []
    errors = []
    for index, item in enumerate(raw_pdfs, start=1):
        slot = int(item.get("index") or index)
        final = _pdf_target(opju, slot, str(item.get("table") or item.get("graph") or "graph"))
        if not item.get("ok") or not item.get("staged"):
            errors.append(_pdf_error(item, final, str(item.get("error") or "PDF export failed")))
            continue
        try:
            info = validate_pdf_file(Path(item["staged"]))
            _install_staged_file(Path(item["staged"]), final, overwrite=bool(job["overwrite"]))
            checked = validate_pdf_file(final)
            installed.append({
                "path": str(final),
                "sha256": _sha256_file(final),
                "size_bytes": final.stat().st_size,
                "pages": checked["pages"],
                "graph": item.get("graph") or "",
                "table": item.get("table") or "",
                "short_name": item.get("short_name") or "",
                "index": slot,
                "header": info["header"],
            })
        except Exception as exc:
            errors.append(_pdf_error(item, final, str(exc) or exc.__class__.__name__))
    if errors or len(installed) != graph_count:
        return installed, "failed", errors
    return installed, "ok", []


def _pdf_error(item: Mapping[str, Any], path: Path, message: str) -> dict[str, Any]:
    return {
        "table": str(item.get("table") or ""),
        "graph": str(item.get("graph") or ""),
        "short_name": str(item.get("short_name") or ""),
        "path": str(path),
        "message": message,
    }


def _graph_identities(raw_pdfs: list[dict[str, Any]]) -> list[dict[str, Any]]:
    identities = []
    for index, item in enumerate(raw_pdfs, start=1):
        identities.append({
            "index": int(item.get("index") or index),
            "table": str(item.get("table") or ""),
            "graph": str(item.get("graph") or ""),
            "short_name": str(item.get("short_name") or ""),
        })
    return identities


def _execute_unified(job, *, progress, cancel_event) -> dict[str, Any]:
    """Legacy single-project export. It holds every selected table until that one file is written."""
    from .exporter import execute_import
    from .planning import create_plan, prepare_plan

    warnings = [
        "unified layout keeps every selected table in memory until the single output is written"
    ]
    record_path = Path(job["record_path"])
    output_dir = Path(job["output_dir"])
    output_dir.mkdir(parents=True, exist_ok=True)
    paths, excluded = _input_paths(job)
    warnings.extend(excluded)
    record, record_warning = _open_record(record_path, resume=job["resume"], retry_ids=job["retry_task_ids"])
    if record_warning:
        warnings.append(record_warning)
    output = output_dir / _with_suffix(str(job["output_name"]), job["format"])
    task = {
        "task_id": "unified",
        "source": str(paths[0]) if paths else "",
        "sources": [str(path) for path in paths],
        "output": str(output),
        "read_options": job["read_options"],
        "plot": job["plot"],
        "pdf_requested": bool(job["pdf"]),
    }
    state = {
        "counts": {name: 0 for name in ("succeeded", "skipped", "failed", "cancelled", "blocked", "pdf_failed")},
        "origin_starts": 0,
        "origin_stops": 0,
        "owned_pids": [],
        "owned_processes": [],
        "recycled": 0,
        "cancelled": False,
    }
    results: list[dict[str, Any]] = []
    if _cancelled(cancel_event):
        state["cancelled"] = True
        _finish_task(record, record_path, results, state, task, 0, 1, progress, warnings,
                     status="cancelled", error=_error("Cancelled", "cancelled before the unified export started"))
        return _result(job, results, state, warnings, len(paths))
    if not paths:
        _finish_task(record, record_path, results, state, task, 0, 1, progress, warnings,
                     status="failed", error=_error("DataImportError", "No supported input files were found"))
        return _result(job, results, state, warnings, 0)
    launches_before = OriginSession.launch_count
    closes_before = OriginSession.close_count
    try:
        plan = create_plan(
            paths,
            output,
            format=job["format"],
            options=job["read_options"],
            overwrite=bool(job["overwrite"]),
            keep_open=bool(job["keep_open"]),
        )
        _apply_unified_plot(plan, job["plot"])
        prepared = prepare_plan(plan)
        receipt = execute_import(prepared)
    except FileExistsError as exc:
        _finish_task(record, record_path, results, state, task, 0, 1, progress, warnings,
                     status="blocked", error=_error(type(exc).__name__, str(exc)))
    except Exception as exc:
        _finish_task(record, record_path, results, state, task, 0, 1, progress, warnings,
                     status="failed", error=_error(type(exc).__name__, str(exc)))
    else:
        _finish_task(
            record, record_path, results, state, task, 0, 1, progress, warnings,
            status="succeeded",
            output=receipt["output"],
            pdf_status="not_applicable",
            source_sha256=_combined_source_hash(paths),
            warnings_for_task=list(receipt.get("warnings") or []),
        )
    state["origin_starts"] = OriginSession.launch_count - launches_before
    state["origin_stops"] = OriginSession.close_count - closes_before
    return _result(job, results, state, warnings, len(paths))


def _apply_unified_plot(plan: dict[str, Any], plot_config: Mapping[str, Any]) -> None:
    """Apply explicit selectors. ``auto`` keeps the plan suggestion for that field."""
    for entry in plan["tables"]:
        kind = plot_config.get("kind", "auto")
        if kind == "none":
            entry["plot"] = {
                "kind": "none", "x": None, "y": [], "y_error": {},
                "title": "" if plot_config.get("title", "auto") == "auto" else plot_config.get("title", ""),
                "x_label": "" if plot_config.get("x_label", "auto") == "auto" else plot_config.get("x_label", ""),
                "y_label": "" if plot_config.get("y_label", "auto") == "auto" else plot_config.get("y_label", ""),
            }
            continue
        if kind != "auto":
            entry["plot"]["kind"] = kind
        for key in ("x", "y", "y_error", "title", "x_label", "y_label"):
            if key in plot_config and plot_config[key] != "auto":
                entry["plot"][key] = plot_config[key]


def _discover_tasks(job: dict[str, Any]) -> tuple[list[dict[str, Any]], list[str]]:
    paths, notes = _input_paths(job)
    tasks = []
    for path in paths:
        override = _override_for(path, job["task_overrides"])
        read_options = _merge_options(job["read_options"], override.get("read_options"))
        plot = _merge_plot(job["plot"], override.get("plot"))
        output = stable_output_path(path, Path(job["output_dir"]), _FORMAT_SUFFIX[job["format"]])
        tasks.append({
            "task_id": task_id_for(path),
            "source": str(path),
            "output": str(output),
            "read_options": read_options,
            "plot": plot,
            "pdf_requested": bool(job["pdf"]),
        })
    return tasks, notes


def _input_paths(job: dict[str, Any]) -> tuple[list[Path], list[str]]:
    from .readers import discover_files

    raw_inputs = [Path(item) for item in job["inputs"]]
    explicit_files = {path.resolve() for path in raw_inputs if path.exists() and path.is_file()}
    discovered = discover_files(raw_inputs)
    output_dir = Path(job["output_dir"]).resolve()
    record_path = Path(job["record_path"]).resolve()
    log_path = log_path_for(record_path).resolve()
    scanned_dirs = [path.resolve() for path in raw_inputs if path.exists() and path.resolve().is_dir()]
    unified_output = (output_dir / _with_suffix(str(job["output_name"]), job["format"])).resolve()
    notes: list[str] = []
    kept: list[Path] = []
    output_inside_scan = any(_is_inside(output_dir, root) and output_dir != root for root in scanned_dirs)
    for path in discovered:
        resolved = path.resolve()
        reason = _exclusion_reason(
            resolved,
            output_dir=output_dir,
            output_inside_scan=output_inside_scan,
            record_path=record_path,
            log_path=log_path,
            unified_output=unified_output,
            explicit=resolved in explicit_files,
        )
        if reason:
            notes.append(f"excluded {resolved}: {reason}")
            continue
        kept.append(resolved)
    if not kept:
        raise DataImportError("No supported input files were found after excluding batch outputs")
    return kept, notes


def _exclusion_reason(
    path: Path,
    *,
    output_dir: Path,
    output_inside_scan: bool,
    record_path: Path,
    log_path: Path,
    unified_output: Path,
    explicit: bool = False,
) -> str | None:
    if path == record_path or path == log_path or path == Path(str(record_path) + ".bak"):
        return "batch record or log"
    if path.name.endswith(".previous") and path.stem == record_path.name:
        return "archived batch record"
    if _GENERATED_NAME.search(path.name) and not explicit:
        return "generated batch output name"
    if path == unified_output:
        return "unified output path"
    if output_inside_scan and _is_inside(path, output_dir):
        return "file is inside the output directory"
    return None


def stable_output_path(source: Path, output_dir: Path, suffix: str) -> Path:
    """Stable file name from the source stem plus a hash of its full path.

    The hash keeps Chinese names readable, cuts very long stems, and separates
    identical file names that live in different directories. It does not depend
    on discovery order.
    """
    stem = _INVALID_CHARS.sub("_", source.stem).strip(" .") or "input"
    ident = hashlib.sha256(os.path.normcase(str(source.resolve())).encode("utf-8")).hexdigest()[:12]
    return Path(output_dir) / f"{stem[:160]}__{ident}{suffix}"


def task_id_for(path: Path) -> str:
    return hashlib.sha256(os.path.normcase(str(path.resolve())).encode("utf-8")).hexdigest()[:16]


def _pdf_target(opju: Path, index: int, table_name: str) -> Path:
    safe = _INVALID_CHARS.sub("_", table_name).strip(" .") or "graph"
    return opju.with_name(f"{opju.stem}__g{index:02d}_{safe[:40]}.pdf")


def _prepare_tables(tables, output: Path, fmt: str, overwrite: bool, plot_config: Mapping[str, Any]) -> PreparedImport:
    from .planning import _unique_target_name, _validate_plot

    planned: list[PlannedTable] = []
    warnings: list[str] = []
    used: set[str] = set()
    for table in tables:
        name = _unique_target_name(table.name, used)
        plot_dict, plot_warnings = _resolve_plot(table, plot_config)
        plot_warnings = list(plot_warnings)
        plot = _validate_plot(plot_dict, table, name, plot_warnings)
        prepared_table = replace(table, warnings=tuple(table.warnings) + tuple(plot_warnings))
        planned.append(PlannedTable(prepared_table, name.strip(), plot))
        warnings.extend(plot_warnings)
    return PreparedImport(tuple(planned), output, fmt, overwrite, False, tuple(dict.fromkeys(warnings)))


def _resolve_plot(table: DataTable, config: Mapping[str, Any]) -> tuple[dict[str, Any], list[str]]:
    from .planning import _suggest_plot

    suggested, warnings = _suggest_plot(table)
    kind = suggested["kind"] if config.get("kind", "auto") == "auto" else config["kind"]
    if kind == "none":
        return {
            "kind": "none",
            "x": None,
            "y": [],
            "y_error": {},
            "title": "" if config.get("title", "auto") == "auto" else config.get("title", ""),
            "x_label": "" if config.get("x_label", "auto") == "auto" else config.get("x_label", ""),
            "y_label": "" if config.get("y_label", "auto") == "auto" else config.get("y_label", ""),
        }, warnings

    def pick(key: str, fallback: Any) -> Any:
        if key not in config or config[key] == "auto":
            return fallback
        return config[key]

    no_suggestion = suggested["kind"] == "none"
    return {
        "kind": kind,
        "x": pick("x", None if no_suggestion else suggested["x"]),
        "y": pick("y", [] if no_suggestion else suggested["y"]),
        "y_error": pick("y_error", {} if no_suggestion else suggested["y_error"]),
        "title": pick("title", suggested.get("title") or table.name),
        "x_label": pick("x_label", "" if no_suggestion else suggested.get("x_label") or ""),
        "y_label": pick("y_label", "" if no_suggestion else suggested.get("y_label") or ""),
    }, warnings


def _plot_public(planned: PlannedTable) -> dict[str, Any]:
    table = planned.table
    return {
        "table": planned.name,
        "kind": planned.plot.kind,
        "x": planned.plot.x,
        "y": list(planned.plot.y),
        "y_error": {table.columns[y].name: table.columns[error].name for y, error in planned.plot.y_error},
        "title": planned.plot.title,
        "x_label": planned.plot.x_label,
        "y_label": planned.plot.y_label,
    }


def _load_tables(path: Path, options: Mapping[str, Any]) -> tuple[list[DataTable], str]:
    from .readers import read_tables

    tables = read_tables(path, **dict(options))
    if not tables:
        raise DataImportError(f"Reader returned no tables for {path}")
    digest = _sha256_file(path)
    for table in tables:
        if str(table.source_hash).lower() != digest:
            raise DataImportError(f"Reader hash did not match the file bytes: {path}")
    return list(tables), digest


def _task_decision(record: Mapping[str, Any], task: Mapping[str, Any], job: Mapping[str, Any]) -> dict[str, Any]:
    """Choose skip, recover, block, or a full run for one recorded task."""
    stored = (record.get("tasks") or {}).get(task["task_id"])
    run = {"kind": "run", "reason": "no record", "keep_pdfs": [], "export": [], "stored": None}
    if not isinstance(stored, dict):
        return run
    if stored.get("status") != "succeeded":
        return {**run, "reason": f"recorded status is {stored.get('status')}", "stored": stored}
    source = Path(task["source"])
    if not source.is_file():
        return {**run, "reason": "source missing", "stored": stored}
    try:
        current_hash = _sha256_file(source)
    except OSError as exc:
        return {**run, "reason": f"source unreadable: {exc}", "stored": stored}
    if current_hash != stored.get("source_sha256"):
        return {**run, "reason": "source bytes changed", "stored": stored}
    if _canonical(stored.get("read_options")) != _canonical(task["read_options"]):
        return {**run, "reason": "read options changed", "stored": stored}
    if _canonical(stored.get("plot")) != _canonical(task["plot"]):
        return {**run, "reason": "plot config changed", "stored": stored}
    if stored.get("format") != job["format"] or bool(stored.get("pdf")) != bool(job["pdf"]):
        return {**run, "reason": "output format or pdf setting changed", "stored": stored}
    fingerprint = _fingerprint(current_hash, task["read_options"], task["plot"], job["format"], job["pdf"], task["output"])
    if stored.get("fingerprint") != fingerprint:
        return {**run, "reason": "task fingerprint changed", "stored": stored}
    output = stored.get("output") or {}
    if str(output.get("path") or "") != str(task["output"]):
        return {**run, "reason": "output path changed", "stored": stored}
    target = Path(str(output.get("path") or ""))
    if not target.is_file() or target.stat().st_size <= 0:
        return {**run, "reason": "output missing or empty", "stored": stored}
    if _sha256_file(target) != output.get("sha256"):
        return {**run, "reason": "output bytes changed", "stored": stored}
    if not job["pdf"] or stored.get("pdf_status") == "not_applicable":
        return {"kind": "skip", "reason": "", "keep_pdfs": [], "export": [], "stored": stored}
    return _pdf_recovery_plan(stored, job, target)


def _pdf_recovery_plan(stored: Mapping[str, Any], job: Mapping[str, Any], opju: Path) -> dict[str, Any]:
    from .session import validate_pdf_file

    expected = _expected_graphs(stored)
    if not expected:
        return _preserved_block(stored, "PDF recovery needs recorded graph identities and this record has none")
    recorded = [item for item in (stored.get("pdfs") or []) if isinstance(item, dict)]
    keep: list[dict[str, Any]] = []
    export: list[dict[str, Any]] = []
    for graph in expected:
        final = _pdf_target(opju, int(graph["index"]), graph["table"])
        receipt = _receipt_for(recorded, final, graph)
        if final.is_file() and final.stat().st_size > 0:
            digest = _sha256_file(final)
            try:
                validate_pdf_file(final)
                valid = True
            except Exception:
                valid = False
            if receipt and receipt.get("sha256") == digest and valid:
                keep.append(_kept_pdf(receipt, graph))
                continue
            if job["overwrite"]:
                export.append({**graph, "path": str(final)})
                continue
            if receipt:
                return _preserved_block(stored, f"pdf bytes changed: {final}")
            return _preserved_block(stored, f"unexpected pdf at {final}; refusing to replace it without overwrite")
        export.append({**graph, "path": str(final)})
    if not export:
        return {"kind": "skip", "reason": "", "keep_pdfs": keep, "export": [], "stored": stored}
    return {"kind": "recover", "reason": "pdf recovery", "keep_pdfs": keep, "export": export, "stored": stored}


def _preserved_block(stored: Mapping[str, Any], reason: str) -> dict[str, Any]:
    return {"kind": "block", "reason": reason, "keep_pdfs": [], "export": [], "stored": stored}


def _expected_graphs(stored: Mapping[str, Any]) -> list[dict[str, Any]]:
    identities = stored.get("graph_identities")
    if isinstance(identities, list) and identities:
        graphs = []
        for index, item in enumerate(identities, start=1):
            if not isinstance(item, dict):
                continue
            graphs.append({
                "index": int(item.get("index") or index),
                "table": str(item.get("table") or ""),
                "graph": str(item.get("graph") or ""),
                "short_name": str(item.get("short_name") or ""),
            })
        return [item for item in graphs if item["table"] and item["graph"]]
    graphs = []
    index = 0
    for item in stored.get("resolved_plot") or []:
        if not isinstance(item, dict) or item.get("kind") == "none":
            continue
        index += 1
        graphs.append({
            "index": index,
            "table": str(item.get("table") or ""),
            "graph": str(item.get("title") or item.get("table") or ""),
            "short_name": "",
        })
    return [item for item in graphs if item["table"] and item["graph"]]


def _receipt_for(recorded: list[dict[str, Any]], final: Path, graph: Mapping[str, Any]) -> dict[str, Any] | None:
    final_key = os.path.normcase(str(final))
    for item in recorded:
        if os.path.normcase(str(item.get("path") or "")) == final_key:
            return item
    matches = [
        item for item in recorded
        if str(item.get("table") or "") == graph["table"] and str(item.get("graph") or "") == graph["graph"]
    ]
    if len(matches) == 1:
        return matches[0]
    return None


def _kept_pdf(receipt: Mapping[str, Any], graph: Mapping[str, Any]) -> dict[str, Any]:
    kept = dict(receipt)
    kept["index"] = int(graph["index"])
    kept["table"] = graph["table"]
    kept["graph"] = graph["graph"]
    kept["short_name"] = str(graph.get("short_name") or receipt.get("short_name") or "")
    return kept


def _reject_unsupported_modes(job: Mapping[str, Any]) -> None:
    """Reject combinations that would ignore an explicit request."""
    if job.get("layout") == "unified" and job.get("pdf"):
        raise DataImportError(
            "unified layout cannot export PDFs; it is the legacy single-project path. Use layout 'per_file'."
        )
    if job.get("layout") == "unified" and (job.get("resume") or job.get("retry_task_ids")):
        raise DataImportError(
            "unified layout does not support resume or retry; it would re-export the whole project "
            "without skip semantics. Use layout 'per_file'."
        )
    if job.get("retry_task_ids") and not job.get("resume"):
        raise DataImportError(
            "retry_task_ids requires resume=True; passing it with resume=False is rejected instead of ignored"
        )


def _remember_owned(state: dict[str, Any], session: Any) -> None:
    """Record every proved process, not only the latest session.

    ``owned_pids`` is the sorted set of those process ids. ``owned_processes``
    keeps the creation time and image when the session proved them. A later
    process that reuses an id has a different creation time and is a different
    entry. Fake sessions that only expose process ids are stored as pid-only
    entries. This is the history of processes this batch proved, not every
    Origin process that happened to appear.
    """
    identities = list(getattr(session, "owned_identities", None) or [])
    if not identities:
        for pid in getattr(session, "owned_pids", []) or []:
            identities.append({"pid": int(pid)})
    seen = {(item.get("pid"), item.get("creation_filetime")) for item in state["owned_processes"]}
    for item in identities:
        copied = dict(item)
        key = (copied.get("pid"), copied.get("creation_filetime"))
        if key in seen:
            continue
        state["owned_processes"].append(copied)
        seen.add(key)
    state["owned_pids"] = sorted({
        int(item["pid"]) for item in state["owned_processes"] if item.get("pid") is not None
    })


def _report_preserved_block(
    record_path, results, state, task, index, count, progress, warnings, reason, pdf_errors=None,
) -> None:
    """Report a block without replacing the stored success receipt."""
    error = _error("DataImportError", reason)
    results.append({
        "task_id": task["task_id"],
        "source": task.get("source"),
        "sources": list(task.get("sources") or ([task["source"]] if task.get("source") else [])),
        "status": "blocked",
        "output": None,
        "pdfs": [],
        "pdf_status": "not_applicable",
        "pdf_errors": list(pdf_errors or []),
        "error": error,
        "warnings": [],
        "resolved_plot": None,
    })
    state["counts"]["blocked"] = state["counts"].get("blocked", 0) + 1
    _emit(record_path, progress, warnings, {
        "type": "task",
        "task_id": task["task_id"],
        "index": index,
        "count": count,
        "status": "blocked",
        "source": task.get("source"),
        "output": task.get("output"),
        "error": error,
        "pdf_status": "not_applicable",
        "pdf_errors": list(pdf_errors or []),
    })


def _restore_exact(path: Path, payload: bytes) -> bool:
    temporary = path.with_name(path.name + ".restore-tmp")
    try:
        temporary.write_bytes(payload)
        os.replace(temporary, path)
        return hashlib.sha256(path.read_bytes()).hexdigest() == hashlib.sha256(payload).hexdigest()
    except OSError:
        return False
    finally:
        if temporary.exists():
            temporary.unlink(missing_ok=True)


def _pdf_failed_keep(session, since_start, stored, decision, source_hash, message: str) -> dict[str, Any]:
    export = decision.get("export") or []
    target = Path(export[0]["path"]) if export else Path(str((stored.get("output") or {}).get("path") or ""))
    return {
        "status": "succeeded",
        "session": session,
        "since_start": since_start,
        "recycle": False,
        "output": stored.get("output"),
        "pdfs": list(decision.get("keep_pdfs") or []),
        "pdf_status": "failed",
        "pdf_errors": [_pdf_error(
            {"table": "", "graph": "", "short_name": ""},
            target,
            message,
        )],
        "source_sha256": source_hash,
        "warnings": [],
        "resolved_plot": stored.get("resolved_plot"),
        "graph_identities": list(stored.get("graph_identities") or []),
        "error": None,
    }


def _preserve_outcome(session, since_start, reason: str) -> dict[str, Any]:
    return {
        "preserve_record": True,
        "status": "blocked",
        "session": session,
        "since_start": since_start,
        "recycle": False,
        "error": _error("DataImportError", reason),
        "pdf_errors": [],
    }


def _recover_one(job, task, decision, state, session, session_factory, since_start) -> dict[str, Any]:
    """Export missing PDFs from the saved project without rewriting that project."""
    from .exporter import _install_staged_file
    from .session import validate_pdf_file

    stored = decision.get("stored") or {}
    output = Path(task["output"])
    source = Path(task["source"])
    try:
        source_hash = _sha256_file(source)
        snapshot = output.read_bytes()
    except OSError as exc:
        return _preserve_outcome(session, since_start, f"PDF recovery could not read source or project: {exc}")
    opju_hash = hashlib.sha256(snapshot).hexdigest()
    if source_hash != stored.get("source_sha256") or opju_hash != (stored.get("output") or {}).get("sha256"):
        return _preserve_outcome(session, since_start, "source or project bytes changed before PDF recovery")
    started_new = False
    if session is None or not session.healthy():
        if session is not None:
            before_stops = int(getattr(session, "stops", 0))
            try:
                session.close()
            finally:
                state["origin_stops"] += int(getattr(session, "stops", 0)) - before_stops
        session = _make_session(job, session_factory)
        before_starts = int(getattr(session, "starts", 0))
        try:
            session.start()
        except Exception as exc:
            state["origin_starts"] += int(getattr(session, "starts", 0)) - before_starts
            return _pdf_failed_keep(session, since_start, stored, decision, source_hash, str(exc) or exc.__class__.__name__)
        state["origin_starts"] += int(getattr(session, "starts", 0)) - before_starts
        _remember_owned(state, session)
        started_new = True
    if started_new:
        since_start = 0
    stage_dir = Path(tempfile.mkdtemp(prefix=f".{output.stem}_pdf_", dir=str(output.parent)))
    try:
        graphs = []
        for item in decision.get("export") or []:
            dest = stage_dir / f"recover_{int(item['index']):02d}.pdf"
            graphs.append({
                "index": int(item["index"]),
                "table": item["table"],
                "graph": item["graph"],
                "short_name": item.get("short_name") or "",
                "source_name": str(source.resolve()),
                "dest": str(dest),
                "final": item["path"],
            })
        try:
            written = list(session.export_saved_pdfs(output, graphs))
        except Exception as exc:
            try:
                current = hashlib.sha256(output.read_bytes()).hexdigest()
            except OSError:
                return _preserve_outcome(
                    session, since_start, f"PDF recovery could not re-read the saved project: {exc}",
                )
            if current != opju_hash and not _restore_exact(output, snapshot):
                return _preserve_outcome(
                    session, since_start, f"PDF recovery changed the saved project and it could not be restored: {exc}",
                )
            return _pdf_failed_keep(
                session, since_start, stored, decision, source_hash, str(exc) or exc.__class__.__name__,
            )
        if len(written) != len(graphs):
            return _pdf_failed_keep(
                session, since_start, stored, decision, source_hash,
                "PDF recovery did not return one result for each requested graph",
            )
        if hashlib.sha256(output.read_bytes()).hexdigest() != opju_hash:
            if not _restore_exact(output, snapshot):
                return _preserve_outcome(session, since_start, "PDF recovery changed the saved project")
            return _pdf_failed_keep(
                session, since_start, stored, decision, source_hash,
                "PDF recovery changed the saved project; the original bytes were restored",
            )
        installed: list[dict[str, Any]] = []
        errors: list[dict[str, Any]] = []
        for spec, item in zip(graphs, written):
            final = Path(spec["final"])
            if not item.get("ok") or not item.get("staged"):
                errors.append(_pdf_error(spec, final, str(item.get("error") or "PDF export failed")))
                continue
            try:
                validate_pdf_file(Path(item["staged"]))
                _install_staged_file(Path(item["staged"]), final, overwrite=bool(job["overwrite"]))
                checked = validate_pdf_file(final)
            except Exception as exc:
                errors.append(_pdf_error(spec, final, str(exc) or exc.__class__.__name__))
                continue
            installed.append({
                "path": str(final),
                "sha256": _sha256_file(final),
                "size_bytes": final.stat().st_size,
                "pages": checked["pages"],
                "graph": spec["graph"],
                "table": spec["table"],
                "short_name": spec["short_name"],
                "index": spec["index"],
                "header": checked["header"],
            })
        if hashlib.sha256(output.read_bytes()).hexdigest() != opju_hash:
            if not _restore_exact(output, snapshot):
                return _preserve_outcome(session, since_start, "PDF install changed the saved project")
            return _pdf_failed_keep(
                session, since_start, stored, decision, source_hash,
                "PDF install changed the saved project; the original bytes were restored",
            )
        merged = list(decision.get("keep_pdfs") or []) + installed
        merged.sort(key=lambda item: int(item.get("index") or 0))
        expected = len(decision.get("keep_pdfs") or []) + len(graphs)
        pdf_status = "ok" if not errors and len(merged) == expected else "failed"
        return {
            "status": "succeeded",
            "session": session,
            "since_start": since_start,
            "recycle": False,
            "output": stored.get("output"),
            "pdfs": merged,
            "pdf_status": pdf_status,
            "pdf_errors": errors,
            "graph_identities": list(stored.get("graph_identities") or []),
            "source_sha256": source_hash,
            "warnings": [],
            "resolved_plot": stored.get("resolved_plot"),
            "error": None,
        }
    finally:
        shutil.rmtree(stage_dir, ignore_errors=True)


def _fingerprint(source_sha: str, read_options, plot, fmt: str, pdf: bool, output: str) -> str:
    return _digest({
        "source_sha256": source_sha,
        "read_options": read_options,
        "plot": plot,
        "format": fmt,
        "pdf": bool(pdf),
        "output": output,
    })


def _finish_task(
    record, record_path, results, state, task, index, count, progress, warnings,
    *, status, error=None, output=None, pdfs=None, pdf_status="not_applicable",
    source_sha256=None, warnings_for_task=None, resolved_plot=None,
    pdf_errors=None, graph_identities=None,
) -> None:
    entry = {
        "task_id": task["task_id"],
        "source": task.get("source"),
        "sources": list(task.get("sources") or ([task["source"]] if task.get("source") else [])),
        "status": status,
        "output": output,
        "pdfs": list(pdfs or []),
        "pdf_status": pdf_status,
        "pdf_errors": list(pdf_errors or []),
        "error": error,
        "warnings": list(warnings_for_task or []),
        "resolved_plot": resolved_plot,
        "read_options": task.get("read_options"),
        "plot": task.get("plot"),
        "format": None,
        "pdf": None,
        "source_sha256": source_sha256,
        "fingerprint": None,
    }
    results.append({key: value for key, value in entry.items() if key in {
        "task_id", "source", "sources", "status", "output", "pdfs", "pdf_status", "pdf_errors",
        "error", "warnings", "resolved_plot",
    }})
    state["counts"][status] = state["counts"].get(status, 0) + 1
    if status == "succeeded" and pdf_status == "failed":
        state["counts"]["pdf_failed"] += 1
    if status == "skipped":
        _emit(record_path, progress, warnings, {
            "type": "task",
            "task_id": task["task_id"],
            "index": index,
            "count": count,
            "status": status,
            "source": task.get("source"),
            "output": None if not output else output.get("path"),
            "error": error,
            "pdf_status": pdf_status,
            "pdf_errors": list(pdf_errors or []),
        })
        return
    event = {
        "type": "task",
        "task_id": task["task_id"],
        "index": index,
        "count": count,
        "status": status,
        "source": task.get("source"),
        "output": None if not output else output.get("path"),
        "error": error,
        "pdf_status": pdf_status,
        "pdf_errors": list(pdf_errors or []),
    }
    _emit(record_path, progress, warnings, event)
    if status != "succeeded":
        record.setdefault("tasks", {})[task["task_id"]] = {
            "task_id": task["task_id"],
            "source": task.get("source"),
            "status": status,
            "error": error,
            "output": output,
        }
    else:
        fmt = (output or {}).get("format")
        record.setdefault("tasks", {})[task["task_id"]] = {
            "task_id": task["task_id"],
            "source": task.get("source"),
            "sources": entry["sources"],
            "status": "succeeded",
            "source_sha256": source_sha256,
            "read_options": task.get("read_options"),
            "plot": task.get("plot"),
            "format": fmt,
            "pdf": bool(task.get("pdf_requested")),
            "pdf_status": pdf_status,
            "output": output,
            "pdfs": list(pdfs or []),
            "pdf_errors": list(pdf_errors or []),
            "graph_identities": list(graph_identities or []),
            "resolved_plot": resolved_plot,
            "fingerprint": _fingerprint(
                source_sha256 or "",
                task.get("read_options"),
                task.get("plot"),
                fmt or "",
                bool(task.get("pdf_requested")),
                str((output or {}).get("path") or task.get("output") or ""),
            ) if source_sha256 else None,
            "error": None,
        }
    try:
        save_batch_record(record_path, record)
    except Exception as exc:
        message = f"task {status} but the record was not saved: {exc}"
        warnings.append(message)
        results[-1]["status"] = "failed"
        results[-1]["error"] = _error("RecordError", message)
        if status in state["counts"] and state["counts"][status]:
            state["counts"][status] -= 1
        state["counts"]["failed"] += 1
        record["tasks"].pop(task["task_id"], None)


def _result(job, results, state, warnings, max_resident: int) -> dict[str, Any]:
    counts = state["counts"]
    ok = (
        not state["cancelled"]
        and counts["failed"] == 0
        and counts["blocked"] == 0
        and counts["cancelled"] == 0
        and counts["pdf_failed"] == 0
    )
    result = {
        "ok": ok,
        "schema_version": BATCH_SCHEMA_VERSION,
        "cancelled": bool(state["cancelled"]),
        "output_dir": job["output_dir"],
        "layout": job["layout"],
        "format": job["format"],
        "record_path": job["record_path"],
        "log_path": str(log_path_for(job["record_path"])),
        "origin": {
            "starts": state["origin_starts"],
            "stops": state["origin_stops"],
            "owned_pids": list(state["owned_pids"]),
            "owned_processes": list(state.get("owned_processes") or []),
            "recycled": state["recycled"],
        },
        "resident": {"max_full_tasks": max_resident, "limit": job["max_resident_tasks"]},
        "counts": counts,
        "tasks": results,
        "warnings": warnings,
    }
    json.dumps(result, ensure_ascii=False)
    return result


def _open_record(record_path: Path, *, resume: bool, retry_ids: list[str]) -> tuple[dict[str, Any], str | None]:
    warning = None
    if not resume:
        _archive_record(record_path)
        record = _empty_record()
        save_batch_record(record_path, record)
        return load_batch_record(record_path), None
    record = load_batch_record(record_path)
    if record.get("_corrupt"):
        warning = f"batch record was corrupt and could not be recovered: {record.get('_corrupt_error')}"
    elif record.get("_recovered_from_backup"):
        warning = "batch record was restored from the last complete backup"
        save_batch_record(record_path, record)
    if retry_ids:
        retry_batch_tasks(record_path, task_ids=list(retry_ids))
        record = load_batch_record(record_path)
    return record, warning


def _archive_record(path: Path) -> None:
    bak = Path(str(path) + ".bak")
    if bak.exists():
        bak.unlink()
    if path.exists():
        os_replace_archive = path.with_name(path.name + ".previous")
        os.replace(path, os_replace_archive)


def _empty_record() -> dict[str, Any]:
    return {"schema_version": BATCH_SCHEMA_VERSION, "kind": RECORD_KIND, "tasks": {}}


def _read_record_file(path: Path) -> dict[str, Any] | None:
    if not path.is_file():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        return {"_corrupt": True, "_corrupt_error": str(exc)}
    if not isinstance(data, dict) or data.get("kind") != RECORD_KIND:
        return {"_corrupt": True, "_corrupt_error": "record kind is missing or unknown"}
    data.setdefault("tasks", {})
    return data


def _atomic_json(path: Path, payload: Mapping[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    raw = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8") + b"\n"
    temporary = path.with_name(path.name + ".tmp")
    with temporary.open("wb") as stream:
        stream.write(raw)
        stream.flush()
        os.fsync(stream.fileno())
    backup = Path(str(path) + ".bak")
    if path.exists():
        os.replace(path, backup)
    os.replace(temporary, path)


def _append_log(record_path: Path, event: Mapping[str, Any]) -> None:
    path = log_path_for(record_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    line = json.dumps(dict(event), ensure_ascii=False, sort_keys=True) + "\n"
    with path.open("a", encoding="utf-8", newline="\n") as stream:
        stream.write(line)
        stream.flush()
        os.fsync(stream.fileno())


def _emit(record_path, progress, warnings, event: dict[str, Any]) -> None:
    try:
        _append_log(Path(record_path), event)
    except OSError as exc:
        warnings.append(f"batch log write failed: {exc}")
    if progress is None:
        return
    try:
        progress(dict(event))
    except Exception as exc:
        warnings.append(f"progress callback failed: {exc}")


def _checked_request(request: Mapping[str, Any]) -> dict[str, Any]:
    if not isinstance(request, Mapping):
        raise DataImportError("batch request must be a mapping")
    if request.get("schema_version") != BATCH_SCHEMA_VERSION:
        raise DataImportError(f"Unsupported batch schema_version: {request.get('schema_version')!r}")
    required = {
        "inputs", "output_dir", "layout", "format", "read_options", "plot", "task_overrides",
        "overwrite", "pdf", "keep_open", "recycle_every", "prefetch", "max_resident_tasks",
        "origin_timeout_s", "origin_retries", "output_name", "record_path", "resume", "retry_task_ids",
    }
    missing = required - set(request)
    if missing:
        raise DataImportError(f"batch request is missing: {', '.join(sorted(missing))}")
    job = json.loads(json.dumps(request))
    return job


def _make_session(job: Mapping[str, Any], factory: Callable[[], Any] | None):
    if factory is not None:
        return factory()
    return OriginSession(timeout_s=float(job["origin_timeout_s"]))


def _is_session_loss(exc: BaseException) -> bool:
    if isinstance(exc, OriginSessionLost):
        return True
    if isinstance(exc, (DataImportError, FileExistsError, FileNotFoundError, ValueError)):
        return False
    name = type(exc).__name__.casefold()
    text = f"{name}: {exc}".casefold()
    if "com_error" in name or "comerror" in name:
        return True
    return any(token in text for token in ("rpc server", "call was rejected", "server execution failed", "0x800"))


def _failed(exc: BaseException, session, since_start, source_hash=None, resolved=None) -> dict[str, Any]:
    return {
        "status": "failed",
        "error": _error(type(exc).__name__, str(exc) or exc.__class__.__name__),
        "session": session,
        "since_start": since_start,
        "source_sha256": source_hash,
        "resolved_plot": resolved,
        "pdf_status": "not_applicable",
    }


def _blocked(exc: BaseException, session, since_start, source_hash=None, resolved=None) -> dict[str, Any]:
    failed = _failed(exc, session, since_start, source_hash, resolved)
    failed["status"] = "blocked"
    return failed


def _error(kind: str, message: str) -> dict[str, str]:
    return {"type": kind, "message": message}


def _stored_output(record, task_id):
    stored = (record.get("tasks") or {}).get(task_id) or {}
    return stored.get("output")


def _stored_pdfs(record, task_id):
    stored = (record.get("tasks") or {}).get(task_id) or {}
    return list(stored.get("pdfs") or [])


def _stored_pdf_status(record, task_id, pdf_enabled: bool) -> str:
    stored = (record.get("tasks") or {}).get(task_id) or {}
    if not pdf_enabled:
        return "not_applicable"
    return str(stored.get("pdf_status") or "not_applicable")


def _combined_source_hash(paths: list[Path]) -> str:
    digest = hashlib.sha256()
    for path in paths:
        digest.update(os.path.normcase(str(path.resolve())).encode("utf-8"))
        digest.update(b"\0")
        digest.update(_sha256_file(path).encode("ascii"))
    return digest.hexdigest()


def _sha256_file(path: Path) -> str:
    from .exporter import _sha256_file as hash_file

    return hash_file(path)


def _public_override(item: Mapping[str, Any]) -> dict[str, Any]:
    if not isinstance(item, Mapping) or not item.get("source"):
        raise DataImportError("task_overrides entries need a source path")
    override: dict[str, Any] = {"source": str(_absolute(str(item["source"])))}
    if "read_options" in item:
        if not isinstance(item["read_options"], Mapping):
            raise DataImportError("task override read_options must be an object")
        override["read_options"] = dict(item["read_options"])
    if "plot" in item:
        if not isinstance(item["plot"], Mapping):
            raise DataImportError("task override plot must be an object")
        override["plot"] = dict(item["plot"])
    return override


def _override_for(path: Path, overrides: list[dict[str, Any]]) -> dict[str, Any]:
    found: dict[str, Any] = {}
    for item in overrides:
        if Path(item["source"]).resolve() == path.resolve():
            found = item
    return found


def _merge_options(base: Mapping[str, Any] | None, extra: Mapping[str, Any] | None) -> dict[str, Any]:
    merged = dict(base or {})
    merged.update(dict(extra or {}))
    return merged


def _merge_plot(base: Mapping[str, Any], extra: Mapping[str, Any] | None) -> dict[str, Any]:
    merged = dict(base)
    if extra:
        unknown = set(extra) - set(_DEFAULT_PLOT)
        if unknown:
            raise DataImportError(f"Unknown plot field(s): {', '.join(sorted(unknown))}")
        merged.update(dict(extra))
    return merged


def _validate_plot_request(plot: Mapping[str, Any], context: str, *, partial: bool = False) -> None:
    if not isinstance(plot, Mapping):
        raise DataImportError(f"{context} must be an object")
    if not partial:
        missing = set(_DEFAULT_PLOT) - set(plot)
        if missing:
            raise DataImportError(f"{context} is missing: {', '.join(sorted(missing))}")
    unknown = set(plot) - set(_DEFAULT_PLOT)
    if unknown:
        raise DataImportError(f"Unknown {context} field(s): {', '.join(sorted(unknown))}")
    if "kind" in plot and plot["kind"] not in _PLOT_KINDS:
        raise DataImportError(f"{context}.kind is not supported: {plot['kind']!r}")


def _with_suffix(name: str, fmt: str) -> str:
    suffix = _FORMAT_SUFFIX[fmt]
    path = Path(name)
    if path.suffix.casefold() != suffix:
        return str(path.with_suffix(suffix) if path.suffix else Path(path.name + suffix))
    return path.name


def _absolute(path: str | Path) -> Path:
    return Path(path).expanduser().resolve(strict=False)


def _is_inside(path: Path, parent: Path) -> bool:
    try:
        path.resolve().relative_to(parent.resolve())
    except ValueError:
        return False
    return True


def _cancelled(event: Any) -> bool:
    if event is None:
        return False
    probe = getattr(event, "is_set", None)
    return bool(probe()) if callable(probe) else False


def _canonical(value: Any) -> Any:
    if isinstance(value, Mapping):
        return {str(key): _canonical(value[key]) for key in sorted(value, key=lambda item: str(item))}
    if isinstance(value, (list, tuple)):
        return [_canonical(item) for item in value]
    return value


def _digest(payload: Mapping[str, Any]) -> str:
    raw = json.dumps(_canonical(payload), ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()
