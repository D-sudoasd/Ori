"""Per-source GUI summaries.

Each call inspects one path. The cached value is the JSON summary returned by
``planning.inspect_inputs`` (column samples only), never a reader ``DataTable``
and never the raw file bytes.

Identity is the file's content SHA-256 plus the requested read options. Size
and mtime are not part of the key: a same-length rewrite still misses.

Failed reads are not cached, so a later reload can succeed after the file is
fixed. Successful entries keep one slot per path; a new hash or a new option
set replaces that slot.

This is not a batch scheduler. A later batch phase may read these summaries
and must not assume every source inspected successfully. Batch export, cancel,
and manifest handling do not belong here.
"""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
from typing import Any

from . import planning


def content_sha256(path: Path) -> str:
    """Hash file bytes without retaining them."""
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def options_fingerprint(options: dict[str, Any] | None) -> str:
    """Stable identity of the read options that affect a summary."""
    source = options or {}
    payload = {
        "sheet": source.get("sheet"),
        "header": source.get("header", "auto"),
        "skip_rows": source.get("skip_rows", 0),
        "delimiter": source.get("delimiter", "auto"),
        "encoding": source.get("encoding", "auto"),
        "missing_values": list(source.get("missing_values") or [""]),
        "formula_policy": source.get("formula_policy", "cached"),
    }
    return json.dumps(payload, ensure_ascii=False, sort_keys=True, default=str)


class SummaryStore:
    """Small path-keyed cache of inspection summaries."""

    def __init__(self) -> None:
        self.inspect_calls = 0
        self.cache_hits = 0
        self.cache_misses = 0
        self._slots: dict[str, tuple[str, str, list[dict[str, Any]]]] = {}

    def inspect_path(self, path: str | Path, options: dict[str, Any] | None = None) -> dict[str, Any]:
        """Inspect one source. Returns a summary outcome and does not raise."""
        source = Path(path).expanduser()
        requested = dict(options or {})
        fingerprint = options_fingerprint(requested)
        try:
            resolved = source.resolve(strict=False)
        except OSError:
            resolved = source
        key = os.path.normcase(str(resolved))
        outcome: dict[str, Any] = {
            "ok": False,
            "path": str(resolved),
            "tables": [],
            "error": "",
            "reused": False,
            "sha256": "",
            "requested_fingerprint": fingerprint,
        }
        slot = self._slots.get(key)
        if slot is not None:
            try:
                digest = content_sha256(resolved)
            except OSError as exc:
                outcome["error"] = f"无法读取文件：{exc}"
                return outcome
            slot_sha, slot_fingerprint, tables = slot
            if digest == slot_sha and fingerprint == slot_fingerprint:
                self.cache_hits += 1
                outcome.update(ok=True, tables=tables, reused=True, sha256=digest)
                return outcome
        self.cache_misses += 1
        self.inspect_calls += 1
        try:
            # One path only. Passing the whole source list would retain every table.
            result = planning.inspect_inputs([resolved], options=requested)
        except Exception as exc:
            self._slots.pop(key, None)
            outcome["error"] = str(exc) or exc.__class__.__name__
            return outcome
        tables = [table for table in result.get("tables", []) if isinstance(table, dict)]
        if not tables:
            self._slots.pop(key, None)
            outcome["error"] = "没有读到数据表。"
            return outcome
        digest = str((tables[0].get("source") or {}).get("sha256") or "")
        if len(digest) != 64:
            try:
                digest = content_sha256(resolved)
            except OSError as exc:
                self._slots.pop(key, None)
                outcome["error"] = f"无法读取文件：{exc}"
                return outcome
        self._slots[key] = (digest, fingerprint, tables)
        outcome.update(ok=True, tables=tables, reused=False, sha256=digest)
        return outcome
