"""Spawn-safe import worker used by the generic data GUI."""

from __future__ import annotations

from pathlib import Path
from typing import Any


def execute_import_request(plan: dict[str, Any], base_dir: str | Path | None = None) -> dict[str, Any]:
    """Prepare an import plan and export it, returning only pickle-safe data."""
    try:
        from .planning import prepare_plan
        from .exporter import execute_import
        from .source_cache import SourceCache

        prepared = prepare_plan(plan, base_dir=Path(base_dir) if base_dir is not None else None, cache=SourceCache())
        result = execute_import(prepared)
        if not isinstance(result, dict):
            raise TypeError("导出器必须返回结果字典")
        if "ok" not in result:
            raise ValueError("导出器结果缺少 ok 字段")
        return result
    except Exception as exc:
        return {"ok": False, "error": str(exc) or exc.__class__.__name__}


def execute_batch_request(
    request: dict[str, Any],
    cancel_event: Any = None,
    progress: Any = None,
) -> dict[str, Any]:
    """Run a batch request in-process. ``cancel_event`` is any object with ``is_set()``."""
    from .batch import execute_batch

    return execute_batch(request, cancel_event=cancel_event, progress=progress)


def batch_worker(send_conn, request: dict[str, Any], cancel_event: Any = None) -> None:
    """Windows-spawn entry. Pass a multiprocessing.Event as ``cancel_event``.

    Progress events are ``{"type": "progress", "event": {...}}``. The final
    message is ``{"type": "result", "result": {...}}``. The request itself must
    stay JSON-safe; the cancel event is a separate spawn argument.
    """
    def progress(event: dict[str, Any]) -> None:
        try:
            send_conn.send({"type": "progress", "event": event})
        except (BrokenPipeError, EOFError, OSError):
            pass

    try:
        result = execute_batch_request(request, cancel_event=cancel_event, progress=progress)
        send_conn.send({"type": "result", "result": result})
    except BaseException as exc:
        try:
            send_conn.send({"type": "result", "result": {"ok": False, "error": str(exc) or exc.__class__.__name__}})
        except (BrokenPipeError, EOFError, OSError):
            pass
    finally:
        try:
            send_conn.close()
        except OSError:
            pass


def import_worker(send_conn, plan: dict[str, Any], base_dir: str | Path | None = None) -> None:
    """Multiprocessing entry point; keep at module scope for Windows spawn."""
    try:
        send_conn.send({"type": "progress", "text": "正在重新读取数据并检查源文件…"})
        result = execute_import_request(plan, base_dir=base_dir)
        send_conn.send({"type": "result", "result": result})
    except BaseException as exc:
        try:
            send_conn.send({"type": "result", "result": {"ok": False, "error": str(exc)}})
        except (BrokenPipeError, EOFError, OSError):
            pass
    finally:
        try:
            send_conn.close()
        except OSError:
            pass
