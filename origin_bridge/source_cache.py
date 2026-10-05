"""Bounded, process-local reuse of parsed sources and inspection summaries.

Every independent operation hashes fresh file bytes. Nested steps share that
snapshot's identity; the exporter still independently verifies sources before
writing. No mtime shortcut, disk cache, cached failure or cached plan approval.
"""
from __future__ import annotations

import copy
import hashlib
import json
import sys
import threading
from collections import OrderedDict
from contextlib import contextmanager
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Any, Callable

from .models import DataImportError, DataTable
from .readers import SUPPORTED_EXTENSIONS, _normalize_options, _parse_tables


def _fingerprint(options: dict[str, Any]) -> str:
    return json.dumps(options, sort_keys=True, ensure_ascii=False)


def _copy_tables(tables: tuple[DataTable, ...]) -> list[DataTable]:
    # Values/columns are immutable. The read-options dict is the only mutable part.
    return [replace(table, read_options=copy.deepcopy(table.read_options)) for table in tables]


@dataclass
class _Entry:
    aliases: dict[str, tuple[DataTable, ...]]
    size: int
    summaries: dict[str, dict[str, Any]] = field(default_factory=dict)


class SourceCache:
    """An LRU cache owned by one CLI command or one MCP server (32 MiB/8 reads)."""

    def __init__(self, *, max_bytes: int = 32 * 1024 * 1024, max_entries: int = 8):
        self.max_bytes = max_bytes
        self.max_entries = max_entries
        self._entries: OrderedDict[tuple[Path, str, str], _Entry] = OrderedDict()
        self._size = 0
        self._checked: dict[Path, str] = {}
        self._depth = 0
        self._lock = threading.RLock()

    @contextmanager
    def operation(self):
        """One read-only decision; discard verified identities when it finishes."""
        with self._lock:
            self._depth += 1
            try:
                yield self
            finally:
                self._depth -= 1
                if self._depth == 0:
                    self._checked.clear()

    def _discard(self, path: Path, digest: str | None = None):
        for key in list(self._entries):
            if key[0] == path and (digest is None or key[1] != digest):
                self._size -= self._entries.pop(key).size

    def read(self, path: Path, **options: Any) -> list[DataTable]:
        """Read once per operation/path; reuse only content-and-option matches."""
        with self.operation():
            return self._read(path, options)

    def read_sheets(self, path: Path, sheets: tuple[str, ...], options: dict[str, Any]) -> list[DataTable]:
        """Open one workbook for exactly the sheets referenced by a saved plan."""
        with self.operation():
            return self._read(path, dict(options, sheet=None), sheets)

    def _read(
        self, path: Path, options: dict[str, Any], selected_sheets: tuple[str, ...] | None = None
    ) -> list[DataTable]:
        path = path.expanduser().resolve()
        settings = _normalize_options(options)
        fingerprint = _fingerprint(settings) if selected_sheets is None else _fingerprint(
            {**settings, "selected_sheets": list(selected_sheets)}
        )
        if not path.is_file():
            self._discard(path)
            raise DataImportError(f"数据文件不存在或不是文件：{path}")
        if path.suffix.casefold() not in SUPPORTED_EXTENSIONS:
            raise DataImportError(f"不支持的文件类型：{path.name}")
        if settings["sheet"] is not None and path.suffix.casefold() not in {".xlsx", ".xlsm", ".xls"}:
            raise DataImportError("sheet 选项仅适用于 Excel 文件。")
        digest = self._checked.get(path)
        data = None
        if digest is None:
            try:
                data = path.read_bytes()
            except OSError:
                self._discard(path)
                raise
            digest = hashlib.sha256(data).hexdigest()
            self._checked[path] = digest
            self._discard(path, digest)
        for key, entry in reversed(self._entries.items()):
            if key[:2] != (path, digest):
                continue
            matched = entry.aliases.get(fingerprint)
            if matched is None and selected_sheets is not None:
                groups = [entry.aliases.get(_fingerprint(dict(settings, sheet=sheet))) for sheet in selected_sheets]
                if all(groups):
                    matched = tuple(table for group in groups for table in group)
            if matched is not None:
                self._entries.move_to_end(key)
                return _copy_tables(matched)
        if data is None:
            # An evicted/oversized source or a different option needs a fresh parse.
            data = path.read_bytes()
            digest = hashlib.sha256(data).hexdigest()
            self._checked[path] = digest
            self._discard(path, digest)
        tables = tuple(_parse_tables(path, data, settings, digest, selected_sheets))
        size = sum(
            sys.getsizeof(column.values) + sum(sys.getsizeof(value) for value in column.values)
            for table in tables for column in table.columns
        )
        if size <= self.max_bytes and self.max_entries > 0:
            aliases = {fingerprint: tables}
            for table in tables:
                aliases.setdefault(_fingerprint(table.read_options), (table,))
            key = (path, digest, fingerprint)
            entry = _Entry(aliases, size)
            self._entries[key] = entry
            self._size += size
            self._evict()
        return _copy_tables(tables)

    def _evict(self):
        while self._entries and (self._size > self.max_bytes or len(self._entries) > self.max_entries):
            self._size -= self._entries.popitem(last=False)[1].size

    def summary(self, table: DataTable, build: Callable[[], dict[str, Any]]) -> dict[str, Any]:
        """Reuse derived samples/ranges; callers receive independent JSON objects."""
        with self._lock:
            return self._summary(table, build)

    def _summary(self, table: DataTable, build: Callable[[], dict[str, Any]]) -> dict[str, Any]:
        fingerprint = _fingerprint(table.read_options)
        for key, entry in reversed(self._entries.items()):
            if key[:2] == (table.source, table.source_hash) and fingerprint in entry.aliases:
                if fingerprint not in entry.summaries:
                    summary = build()
                    # Bound derived summaries too (including exceptionally long samples).
                    size = len(json.dumps(summary, ensure_ascii=False).encode("utf-8")) * 4
                    entry.summaries[fingerprint] = summary
                    entry.size += size
                    self._size += size
                    self._evict()
                return copy.deepcopy(entry.summaries[fingerprint])
        return build()
