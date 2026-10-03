"""Persistent batch configuration saved beside a core record.

The core record stores task receipts. Failed and cancelled receipts do not
include the global read options or plot request, so a later window cannot
rebuild the batch from the receipt alone. This file keeps the JSON-safe
request, the record path, and the input paths. The GUI does not edit success
receipts inside the core record.
"""

from __future__ import annotations

import json
import os
import uuid
from pathlib import Path
from typing import Any, Mapping

CONFIG_KIND = "ori-batch-config"
CONFIG_SCHEMA_VERSION = 1
CONFIG_FILENAME = "batch-config.json"
_RECORD_PREFIX = "batch-record-"


def config_path_for(record_path: str | Path) -> Path:
    """Return the config file that belongs to this record.

    The legacy pair is ``batch-record.json`` and ``batch-config.json``.
    A fresh window batch uses ``batch-record-<id>.json`` and the config is
    named from that same id, so a later batch in the folder does not replace
    the earlier record or config.
    """
    record = Path(record_path).expanduser().resolve(strict=False)
    name = record.name
    if name == "batch-record.json":
        return record.with_name(CONFIG_FILENAME)
    if name.startswith(_RECORD_PREFIX) and name.endswith(".json"):
        ident = name[len(_RECORD_PREFIX):-len(".json")]
        if ident and ident == Path(ident).name:
            return record.with_name(f"batch-config-{ident}.json")
    return record.with_name(f"{record.stem}.batch-config.json")


def allocate_record_path(output_dir: str | Path) -> Path:
    """Pick a new record path that does not reuse an existing record or config."""
    directory = Path(output_dir).expanduser().resolve(strict=False)
    for _attempt in range(8):
        candidate = directory / f"{_RECORD_PREFIX}{uuid.uuid4().hex}.json"
        if candidate.exists() or config_path_for(candidate).exists():
            continue
        return candidate
    raise FileExistsError("无法为新批次分配不冲突的记录名。")


def save_batch_config(
    path: str | Path,
    *,
    name: str,
    record_path: str | Path,
    inputs: list[str],
    request: Mapping[str, Any],
) -> Path:
    """Atomically write a batch config. Raises if the request is not JSON-safe."""
    target = Path(path).expanduser().resolve(strict=False)
    record = str(Path(record_path).expanduser().resolve(strict=False))
    safe_inputs = [str(Path(item).expanduser().resolve(strict=False)) for item in inputs]
    safe_request = json.loads(json.dumps(request, ensure_ascii=False))
    if not isinstance(safe_request, dict):
        raise ValueError("批次请求必须是 JSON 对象。")
    request_inputs = [str(Path(item).expanduser().resolve(strict=False)) for item in safe_request.get("inputs") or []]
    if _same_paths(safe_inputs, request_inputs) is False:
        raise ValueError("批次配置的输入路径必须与请求里的 inputs 一致。")
    if _norm(safe_request.get("record_path", "")) != _norm(record):
        raise ValueError("批次配置的 record_path 必须与请求里的记录路径一致。")
    payload = {
        "schema_version": CONFIG_SCHEMA_VERSION,
        "kind": CONFIG_KIND,
        "name": str(name or "").strip(),
        "record_path": record,
        "inputs": safe_inputs,
        "request": safe_request,
    }
    _atomic_json(target, payload)
    return target


def load_batch_config(path: str | Path) -> dict[str, Any]:
    """Load a batch config. A core record without this file is not enough."""
    source = Path(path).expanduser()
    try:
        data = json.loads(source.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"无法读取批次配置：{exc}") from exc
    if not isinstance(data, dict) or data.get("kind") != CONFIG_KIND:
        raise ValueError("这不是可恢复的批次配置。只有运行记录时不能还原整批读取和绘图设置。")
    if data.get("schema_version") != CONFIG_SCHEMA_VERSION:
        raise ValueError("批次配置的 schema_version 不受支持。")
    request = data.get("request")
    inputs = data.get("inputs")
    record_path = data.get("record_path")
    if not isinstance(request, dict) or not isinstance(inputs, list) or not record_path:
        raise ValueError("批次配置缺少 request、inputs 或 record_path。")
    if any(not isinstance(item, str) or not item for item in inputs):
        raise ValueError("批次配置的输入路径无效。")
    request_inputs = request.get("inputs")
    if not isinstance(request_inputs, list) or _same_paths(inputs, request_inputs) is False:
        raise ValueError("批次配置的 inputs 与 request.inputs 不一致。")
    if _norm(str(request.get("record_path") or "")) != _norm(str(record_path)):
        raise ValueError("批次配置的 record_path 与 request.record_path 不一致。")
    json.dumps(request, ensure_ascii=False)
    data["name"] = str(data.get("name") or "")
    return data


def find_config_for_record(record_path: str | Path) -> Path | None:
    """Find the config that names this record. A name inside the file is the user label."""
    candidate = config_path_for(record_path)
    if not candidate.is_file():
        return None
    try:
        data = load_batch_config(candidate)
    except ValueError:
        return None
    if _norm(str(data.get("record_path") or "")) != _norm(str(Path(record_path).expanduser().resolve(strict=False))):
        return None
    return candidate


def _same_paths(left: list[Any], right: list[Any]) -> bool:
    if len(left) != len(right):
        return False
    return [_norm(str(item)) for item in left] == [_norm(str(item)) for item in right]


def _norm(path: str) -> str:
    if not path:
        return ""
    return os.path.normcase(str(Path(path).expanduser().resolve(strict=False)))


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
