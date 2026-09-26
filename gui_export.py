"""Spawn-safe export worker used by the Tk GUI."""

from __future__ import annotations

from pathlib import Path
from typing import Any


def execute_export_request(request: dict[str, Any]) -> dict[str, Any]:
    """Run one export and return only pickle-safe paths or an error string."""
    import spectra_to_origin as sto

    try:
        labels = sto.AxisLabels(*request["axis_labels"])
        files = request["files"]
        groups = request["groups"]
        output = Path(request["output"])
        if request["kind"] == "opju":
            if sto.origin_process_running():
                raise sto.OriginExportError(
                    "检测到 Origin Pro 正在运行。请先保存并关闭 Origin，再重新导出；"
                    "为保护当前工程，本次没有启动 Origin。"
                )
            opju_path, xlsx_path, csv_path = sto.export_origin_project(
                files,
                output,
                layout=request["layout"],
                groups=groups,
                show_origin=request["show_origin"],
                keep_open=request["keep_open"],
                also_xlsx=request["also_xlsx"],
                axis_labels=labels,
            )
            paths = {
                "opju": str(opju_path),
                "xlsx": str(xlsx_path) if xlsx_path else None,
                "csv": str(csv_path) if csv_path else None,
            }
        elif request["kind"] == "xlsx":
            xlsx_path, csv_path = sto.export_spectra(
                files,
                output,
                layout=request["layout"],
                groups=groups,
                write_csv_copy=True,
                axis_labels=labels,
            )
            paths = {"opju": None, "xlsx": str(xlsx_path), "csv": str(csv_path) if csv_path else None}
        else:
            raise ValueError(f"未知导出类型：{request['kind']}")

        _verify_output(paths)
        return {"ok": True, "paths": paths}
    except Exception as exc:
        return {"ok": False, "error": str(exc) or exc.__class__.__name__}


def _verify_output(paths: dict[str, str | None]) -> None:
    for label in ("opju", "xlsx", "csv"):
        raw_path = paths[label]
        if raw_path is None:
            continue
        path = Path(raw_path)
        if not path.exists():
            raise OSError(f"导出返回成功，但{label.upper()} 文件不存在：{path}")
        if path.is_file() and path.stat().st_size == 0:
            raise OSError(f"导出返回成功，但{label.upper()} 文件为空：{path}")
        if path.is_dir() and not any(item.is_file() and item.stat().st_size > 0 for item in path.glob("*.csv")):
            raise OSError(f"导出返回成功，但{label.upper()} 目录里没有有效 CSV：{path}")


def export_worker(send_conn, request: dict[str, Any]) -> None:
    """Multiprocessing entry point; keep it at module scope for Windows spawn."""
    try:
        send_conn.send({"type": "progress", "text": "后台正在准备导出…"})
        result = execute_export_request(request)
        send_conn.send({"type": "result", "result": result})
    except BaseException as exc:
        try:
            send_conn.send({"type": "result", "result": {"ok": False, "error": str(exc)}})
        except (BrokenPipeError, EOFError, OSError):
            pass
    finally:
        send_conn.close()
