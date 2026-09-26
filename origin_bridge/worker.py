"""Spawn-safe import worker used by the generic data GUI."""

from __future__ import annotations

from pathlib import Path
from typing import Any


def execute_import_request(plan: dict[str, Any], base_dir: str | Path | None = None) -> dict[str, Any]:
    """Prepare an import plan and export it, returning only pickle-safe data."""
    try:
        from .planning import prepare_plan
        from .exporter import execute_import

        prepared = prepare_plan(plan, base_dir=Path(base_dir) if base_dir is not None else None)
        result = execute_import(prepared)
        if not isinstance(result, dict):
            raise TypeError("导出器必须返回结果字典")
        if "ok" not in result:
            raise ValueError("导出器结果缺少 ok 字段")
        return result
    except Exception as exc:
        return {"ok": False, "error": str(exc) or exc.__class__.__name__}


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
