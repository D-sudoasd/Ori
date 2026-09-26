"""Shared immutable data contracts; numeric text is retained until export."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


class DataImportError(ValueError):
    """A source file or import plan cannot be interpreted without data loss."""


@dataclass(frozen=True)
class DataColumn:
    name: str
    kind: str  # number, text, datetime
    values: tuple[str | None, ...]
    unit: str = ""
    original_name: str = ""


@dataclass(frozen=True)
class DataTable:
    source: Path
    source_sheet: str | None
    name: str
    columns: tuple[DataColumn, ...]
    source_hash: str
    read_options: dict[str, Any] = field(default_factory=dict)
    warnings: tuple[str, ...] = ()

    @property
    def n_rows(self) -> int:
        return len(self.columns[0].values) if self.columns else 0


@dataclass(frozen=True)
class PlotSpec:
    kind: str = "none"  # none, line, scatter, line_symbol, column
    x: int | None = None  # None uses row number; indices are zero based
    y: tuple[int, ...] = ()
    y_error: tuple[tuple[int, int], ...] = ()  # (Y index, error index)
    title: str = ""
    x_label: str = ""
    y_label: str = ""


@dataclass(frozen=True)
class PlannedTable:
    table: DataTable
    name: str
    plot: PlotSpec


@dataclass(frozen=True)
class PreparedImport:
    tables: tuple[PlannedTable, ...]
    output: Path
    format: str
    overwrite: bool = False
    keep_open: bool = False
    warnings: tuple[str, ...] = ()
