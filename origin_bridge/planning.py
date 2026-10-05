"""Versioned, deterministic inspect and import-plan contracts for Origin."""

from __future__ import annotations

import json
import re
from contextlib import nullcontext
from copy import deepcopy
from dataclasses import replace
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import TYPE_CHECKING, Any, Mapping

if TYPE_CHECKING:
    from .source_cache import SourceCache

from .models import (
    DataColumn,
    DataImportError,
    DataTable,
    PlannedTable,
    PlotSpec,
    PreparedImport,
)


SCHEMA_VERSION = 1
_READ_OPTION_KEYS = frozenset(
    {"sheet", "header", "skip_rows", "delimiter", "encoding", "missing_values", "formula_policy"}
)
_READ_DELIMITERS = frozenset({"auto", "whitespace", ",", ";", "\t"})
_PLOT_KINDS = frozenset({"none", "line", "scatter", "line_symbol", "column"})
_FORMAT_SUFFIX = {"opju": ".opju", "xlsx": ".xlsx"}
_SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
_ERROR_COLUMN_RE = re.compile(
    r"^(?:err(?:or)?|uncertainty|std|sd|sem|standard\s+(?:deviation|error)|误差|标准差|标准误)$",
    re.IGNORECASE,
)


def _discover_files(paths: list[Path]) -> list[Path]:
    from .readers import discover_files

    return discover_files(paths)


def _read_tables(path: Path, **options: Any) -> list[DataTable]:
    from .readers import read_tables

    return read_tables(path, **options)


def _absolute(path: Path | str, base_dir: Path | None = None) -> Path:
    candidate = Path(path).expanduser()
    if not candidate.is_absolute():
        candidate = (base_dir if base_dir is not None else Path.cwd()) / candidate
    return candidate.resolve(strict=False)


def _source_identity(table: DataTable, requested_path: Path) -> tuple[Path, str]:
    source = _absolute(table.source)
    if source != requested_path:
        raise DataImportError(
            f"Reader returned a different source path for {requested_path}: {source}"
        )
    digest = str(table.source_hash).lower()
    if not _SHA256_RE.fullmatch(digest):
        raise DataImportError(f"Reader returned an invalid SHA-256 for {source}")
    return source, digest


def _validate_options(options: Any, *, replayable: bool) -> dict[str, Any]:
    if not isinstance(options, Mapping):
        raise DataImportError("source.options must be a JSON object")
    unknown = set(options) - _READ_OPTION_KEYS
    if unknown:
        raise DataImportError(f"Unknown source.options field(s): {', '.join(sorted(map(str, unknown)))}")
    if replayable and set(options) != _READ_OPTION_KEYS:
        missing = _READ_OPTION_KEYS - set(options)
        raise DataImportError(f"source.options is missing: {', '.join(sorted(missing))}")

    result = dict(options)
    if "sheet" in result:
        sheet = result["sheet"]
        if sheet is not None and (not isinstance(sheet, str) or not sheet.strip()):
            raise DataImportError("source.options.sheet must be null or a non-empty sheet name")
    if "header" in result:
        header = result["header"]
        if replayable:
            if not isinstance(header, bool):
                raise DataImportError("source.options.header must be a resolved boolean")
        elif not isinstance(header, bool) and header != "auto":
            raise DataImportError("options.header must be true, false, or 'auto'")
    if "skip_rows" in result:
        skip_rows = result["skip_rows"]
        if isinstance(skip_rows, bool) or not isinstance(skip_rows, int) or skip_rows < 0:
            raise DataImportError("source.options.skip_rows must be a non-negative integer")
    if "delimiter" in result and (
        not isinstance(result["delimiter"], str) or result["delimiter"] not in _READ_DELIMITERS
    ):
        raise DataImportError("source.options.delimiter must be auto, whitespace, comma, semicolon, or tab")
    if "encoding" in result and (not isinstance(result["encoding"], str) or not result["encoding"].strip()):
        raise DataImportError("source.options.encoding must be a non-empty string")
    if "missing_values" in result:
        values = result["missing_values"]
        if not isinstance(values, list) or any(not isinstance(value, str) for value in values):
            raise DataImportError("source.options.missing_values must be an array of strings")
    if "formula_policy" in result and (
        not isinstance(result["formula_policy"], str) or result["formula_policy"] not in {"cached", "text"}
    ):
        raise DataImportError("source.options.formula_policy must be 'cached' or 'text'")
    return result


def _validate_table_shape(table: DataTable) -> None:
    if not isinstance(table, DataTable):
        raise DataImportError("Reader must return DataTable objects")
    if not isinstance(table.name, str) or not table.name.strip():
        raise DataImportError(f"Reader returned a table without a valid name: {table.source}")
    if not table.columns:
        raise DataImportError(f"Table has no columns: {table.source}")
    row_count = table.n_rows
    seen: set[str] = set()
    for index, column in enumerate(table.columns):
        if not isinstance(column, DataColumn):
            raise DataImportError(f"Column {index} in {table.source} is not a DataColumn")
        if not isinstance(column.name, str) or not column.name.strip():
            raise DataImportError(f"Column {index} in {table.source} has no name")
        if column.name.casefold() in seen:
            raise DataImportError(f"Duplicate reader column name {column.name!r} in {table.source}")
        seen.add(column.name.casefold())
        if not isinstance(column.kind, str) or column.kind not in {"number", "text", "datetime"}:
            raise DataImportError(f"Unsupported column kind {column.kind!r} in {table.source}")
        if not isinstance(column.unit, str):
            raise DataImportError(f"Column unit must be a string: {column.name}")
        if len(column.values) != row_count:
            raise DataImportError(f"Column lengths differ in table {table.name!r}")
        if any(value is not None and not isinstance(value, str) for value in column.values):
            raise DataImportError(f"Column values must be strings or null: {column.name}")


def _numeric_decimal(value: str, *, context: str) -> Decimal:
    try:
        number = Decimal(value.strip().replace("d", "e").replace("D", "E"))
    except (InvalidOperation, AttributeError) as exc:
        raise DataImportError(f"Expected a numeric value in {context}: {value!r}") from exc
    if not number.is_finite():
        raise DataImportError(f"Expected a finite numeric value in {context}: {value!r}")
    return number


def _column_range(column: DataColumn) -> dict[str, str | None]:
    present = [value for value in column.values if value is not None]
    if column.kind == "number":
        if not present:
            return {"min": None, "max": None}
        values = [_numeric_decimal(value, context=column.name) for value in present]
        return {"min": str(min(values)), "max": str(max(values))}
    if column.kind == "datetime" and present:
        def utc_key(value: str) -> datetime:
            try:
                from .dates import parse_iso_datetime
                parsed = parse_iso_datetime(value)
            except ValueError as exc:
                raise DataImportError(f"Invalid ISO datetime in column {column.name!r}: {value!r}") from exc
            if parsed.tzinfo is None:
                parsed = parsed.replace(tzinfo=timezone.utc)
            return parsed.astimezone(timezone.utc)

        keyed = [(utc_key(value), value) for value in present]
        return {"min": min(keyed, key=lambda item: item[0])[1], "max": max(keyed, key=lambda item: item[0])[1]}
    return {"min": None, "max": None}


def _unique_target_name(name: str, used: set[str]) -> str:
    base = name.strip()
    candidate = base
    suffix = 2
    while candidate.casefold() in used:
        candidate = f"{base}_{suffix}"
        suffix += 1
    used.add(candidate.casefold())
    return candidate


def _plot_dict(
    kind: str,
    x: int | None,
    y: list[int],
    *,
    title: str,
    x_label: str,
    y_label: str,
    y_error: dict[str, int] | None = None,
) -> dict[str, Any]:
    return {
        "kind": kind,
        "x": x,
        "y": y,
        "y_error": y_error or {},
        "title": title,
        "x_label": x_label,
        "y_label": y_label,
    }


def _suggest_plot(table: DataTable) -> tuple[dict[str, Any], list[str]]:
    numeric = [index for index, column in enumerate(table.columns) if column.kind == "number"]
    dates = [index for index, column in enumerate(table.columns) if column.kind == "datetime"]
    text = [index for index, column in enumerate(table.columns) if column.kind == "text"]
    warnings: list[str] = []

    if text and text[0] == 0 and numeric:
        category_column = table.columns[0]
        categories = {value for value in category_column.values if value is not None}
        if len(categories) > 50:
            warnings.append(
                f"Column {category_column.name!r} has {len(categories)} categories; a column plot may be crowded."
            )
        y_indices, error_mapping = _suggest_uncertainties(table, numeric, warnings)
        return (
            _plot_dict(
                "column",
                0,
                y_indices,
                title=table.name,
                x_label=category_column.name,
                y_label=_suggest_y_label(table, y_indices, warnings),
                y_error=error_mapping,
            ),
            warnings,
        )

    if not dates and len(numeric) == 2:
        error_indices = [index for index in numeric if _is_error_column(table.columns[index])]
        if len(error_indices) == 1:
            signal_index = next(index for index in numeric if index not in error_indices)
            _, error_mapping = _suggest_uncertainties(table, numeric, warnings)
            return (
                _plot_dict("line", None, [signal_index], title=table.name, x_label="Row",
                           y_label=_suggest_y_label(table, [signal_index], warnings),
                           y_error=error_mapping),
                warnings,
            )

    x_candidates = dates + numeric
    if x_candidates:
        x_index = x_candidates[0]
        y_indices, error_mapping = _suggest_uncertainties(
            table, [index for index in numeric if index != x_index], warnings
        )
        if y_indices:
            x_column = table.columns[x_index]
            return (
                _plot_dict(
                    "line",
                    x_index,
                    y_indices,
                    title=table.name,
                    x_label=_axis_label(x_column.name, x_column.unit),
                    y_label=_suggest_y_label(table, y_indices, warnings),
                    y_error=error_mapping,
                ),
                warnings,
            )
        if len(numeric) == 1:
            y_index = numeric[0]
            y_column = table.columns[y_index]
            return (
                _plot_dict(
                    "line",
                    None,
                    [y_index],
                    title=table.name,
                    x_label="Row",
                    y_label=_axis_label(y_column.name, y_column.unit),
                ),
                warnings,
            )

    return (_plot_dict("none", None, [], title=table.name, x_label="", y_label=""), warnings)


def _is_error_column(column: DataColumn) -> bool:
    name = column.name.strip()
    # Units commonly appear as a parenthetical suffix, e.g. Error(MPa).
    name = re.sub(r"\s*[([][^\])]*[)\]]\s*$", "", name).strip()
    return bool(_ERROR_COLUMN_RE.fullmatch(name))


def _suggest_uncertainties(
    table: DataTable,
    y_indices: list[int],
    warnings: list[str],
) -> tuple[list[int], dict[str, int]]:
    possible_errors = [index for index in y_indices if _is_error_column(table.columns[index])]
    signal_indices = [index for index in y_indices if index not in possible_errors]
    if len(possible_errors) == 1 and len(signal_indices) == 1:
        error_index = possible_errors[0]
        error_column = table.columns[error_index]
        y_column = table.columns[signal_indices[0]]
        same_units = error_column.unit.strip() == y_column.unit.strip()
        try:
            valid = all(
                value is None or _numeric_decimal(value, context=error_column.name) >= 0
                for value in error_column.values
            )
        except DataImportError:
            valid = False
        if valid and same_units:
            y_index = signal_indices[0]
            return signal_indices, {table.columns[y_index].name: error_index}
        if valid and not same_units:
            warnings.append(
                f"Possible uncertainty column {error_column.name!r} has unit {error_column.unit!r}, "
                f"which does not match Y column {y_column.name!r} unit {y_column.unit!r}; "
                "it was left out of the suggested plot and was not assigned as Y error."
            )
            return signal_indices, {}
        warnings.append(
            f"Possible uncertainty column {error_column.name!r} contains invalid or negative values; "
            "it was left out of the suggested plot and was not assigned as Y error."
        )
        return signal_indices, {}
    if possible_errors and signal_indices:
        names = ", ".join(table.columns[index].name for index in possible_errors)
        warnings.append(
            f"Possible uncertainty column(s) {names} were omitted from the suggested plot because "
            "their Y pairing is ambiguous; assign plot.y_error explicitly if appropriate."
        )
        return signal_indices, {}
    return y_indices, {}


def _suggest_y_label(table: DataTable, y_indices: list[int], warnings: list[str]) -> str:
    if len(y_indices) == 1:
        column = table.columns[y_indices[0]]
        return _axis_label(column.name, column.unit)
    units = {table.columns[index].unit.strip() for index in y_indices}
    if len(units) > 1:
        warnings.append("Selected Y columns use mixed units on one axis; their values are not normalized.")
        return "Value"
    unit = next(iter(units), "")
    return f"Value ({unit})" if unit else "Value"


def _axis_label(name: str, unit: str) -> str:
    return f"{name} ({unit})" if unit else name


def _read_options_for_table(table: DataTable) -> dict[str, Any]:
    options = _validate_options(table.read_options, replayable=True)
    # A workbook's selected sheet is part of a deterministic table identity.
    if table.source_sheet is not None and options["sheet"] is None:
        options["sheet"] = table.source_sheet
    return options


def _read_inputs(
    paths: list[Path], options: dict[str, Any] | None, cache: SourceCache | None = None
) -> list[tuple[Path, DataTable]]:
    if not isinstance(paths, list) or not paths:
        raise DataImportError("At least one input path is required")
    raw_paths: list[Path] = []
    for raw in paths:
        if not isinstance(raw, (str, Path)):
            raise DataImportError("Input paths must be strings or Path objects")
        raw_paths.append(_absolute(raw))
    request_options = _validate_options(options if options is not None else {}, replayable=False)
    discovered = _discover_files(raw_paths)
    if not discovered:
        raise DataImportError("No supported input files were found")
    result: list[tuple[Path, DataTable]] = []
    for raw_path in discovered:
        path = _absolute(raw_path)
        if not path.is_file():
            raise DataImportError(f"Input is not a readable file: {path}")
        tables = (cache.read if cache is not None else _read_tables)(path, **request_options)
        if not isinstance(tables, list) or not tables:
            raise DataImportError(f"Reader returned no tables for {path}")
        for table in tables:
            if not isinstance(table, DataTable):
                raise DataImportError("Reader must return DataTable objects")
            _source_identity(table, path)
            _read_options_for_table(table)
            result.append((path, table))
    return result


def inspect_inputs(
    paths: list[Path], options: dict[str, Any] | None = None, *, cache: SourceCache | None = None
) -> dict[str, Any]:
    """Read input files and return a JSON-safe schema and plot suggestions."""
    tables: list[dict[str, Any]] = []
    used_names: set[str] = set()
    with cache.operation() if cache is not None else nullcontext():
        for path, table in _read_inputs(paths, options, cache):
            def build():
                return _inspect_table(path, table)
            summary = cache.summary(table, build) if cache is not None else build()
            summary["name"] = _unique_target_name(table.name, used_names)
            tables.append(summary)
    return {"schema_version": SCHEMA_VERSION, "tables": tables}


def _inspect_table(path: Path, table: DataTable) -> dict[str, Any]:
    # A cached summary already validated these immutable columns and cell values.
    _validate_table_shape(table)
    source_path, digest = _source_identity(table, path)
    read_options = _read_options_for_table(table)
    plot, suggestion_warnings = _suggest_plot(table)
    columns = [
        {
            "index": index, "name": column.name, "kind": column.kind, "unit": column.unit,
            "missing": sum(value is None for value in column.values),
            "sample": list(column.values[:5]), "range": _column_range(column),
        }
        for index, column in enumerate(table.columns)
    ]
    return {
        "source": {
            "path": str(source_path),
            "sheet": table.source_sheet if table.source_sheet is not None else read_options["sheet"],
            "sha256": digest, "options": read_options,
        },
        "name": table.name, "n_rows": table.n_rows, "columns": columns,
        "warnings": list(table.warnings) + suggestion_warnings, "suggested_plot": plot,
    }


def _output_path(path: Path | str, file_format: str) -> Path:
    if not isinstance(file_format, str) or file_format not in _FORMAT_SUFFIX:
        raise DataImportError("output.format must be 'opju' or 'xlsx'")
    output = _absolute(path)
    suffix = _FORMAT_SUFFIX[file_format]
    if output.suffix.casefold() != suffix:
        output = output.with_suffix(suffix) if output.suffix else output.with_name(output.name + suffix)
    return output


def _default_column_label(column: DataColumn) -> str:
    return _axis_label(column.name, column.unit)


def create_plan(
    paths: list[Path],
    output: Path,
    format: str = "opju",
    options: dict[str, Any] | None = None,
    overwrite: bool = False,
    keep_open: bool = False,
    *,
    cache: SourceCache | None = None,
) -> dict[str, Any]:
    """Build an explicit version-1 import plan without writing files or starting Origin."""
    if not isinstance(overwrite, bool) or not isinstance(keep_open, bool):
        raise DataImportError("overwrite and keep_open must be booleans")
    if not isinstance(format, str) or format not in _FORMAT_SUFFIX:
        raise DataImportError("output.format must be 'opju' or 'xlsx'")
    if not isinstance(output, (str, Path)):
        raise DataImportError("output must be a path string or pathlib.Path")
    if format == "xlsx" and keep_open:
        raise DataImportError("output.keep_open is only valid for opju output")
    inspection = inspect_inputs(paths, options, cache=cache)
    target = _output_path(output, format)
    plan_tables = []
    for table in inspection["tables"]:
        columns = table["columns"]
        plot = dict(table["suggested_plot"])
        x_index = plot["x"]
        y_indices = plot["y"]
        x_label = "Row" if x_index is None else _default_column_label_from_inspection(columns[x_index])
        y_label = _default_column_label_from_inspection(columns[y_indices[0]]) if y_indices else ""
        plot["x_label"] = x_label
        plot["y_label"] = y_label
        source = table["source"]
        plan_tables.append(
            {
                "source": {key: source[key] for key in ("path", "sha256", "options")},
                "name": table["name"],
                "plot": plot,
                "column_labels": {},
            }
        )
    return {
        "schema_version": SCHEMA_VERSION,
        "output": {
            "path": str(target),
            "format": format,
            "overwrite": overwrite,
            "keep_open": keep_open,
        },
        "tables": plan_tables,
    }


def _default_column_label_from_inspection(column: Mapping[str, Any]) -> str:
    return _axis_label(str(column["name"]), str(column["unit"]))


def _unknown_keys(value: Mapping[str, Any], allowed: set[str] | frozenset[str], context: str) -> None:
    unknown = set(value) - set(allowed)
    if unknown:
        raise DataImportError(f"Unknown {context} field(s): {', '.join(sorted(map(str, unknown)))}")


def _require_object(value: Any, context: str, allowed: set[str] | frozenset[str]) -> Mapping[str, Any]:
    if not isinstance(value, Mapping):
        raise DataImportError(f"{context} must be a JSON object")
    _unknown_keys(value, allowed, context)
    return value


def _require_string(value: Any, context: str, *, nonempty: bool = True) -> str:
    if not isinstance(value, str) or (nonempty and not value.strip()):
        raise DataImportError(f"{context} must be {'a non-empty' if nonempty else 'a'} string")
    return value


def _selector(value: Any, columns: tuple[DataColumn, ...], context: str) -> int:
    if isinstance(value, bool):
        raise DataImportError(f"{context} must be a column name or zero-based integer index")
    if isinstance(value, int):
        if value < 0 or value >= len(columns):
            raise DataImportError(f"{context} column index {value} is out of range")
        return value
    if isinstance(value, str):
        matches = [index for index, column in enumerate(columns) if column.name == value]
        if len(matches) == 1:
            return matches[0]
        if not matches:
            raise DataImportError(f"{context} refers to unknown source column {value!r}")
        raise DataImportError(f"{context} column name {value!r} is ambiguous")
    raise DataImportError(f"{context} must be a column name or zero-based integer index")


def _validate_plot(
    value: Any,
    table: DataTable,
    table_name: str,
    warnings: list[str] | None = None,
    label_columns: tuple[DataColumn, ...] | None = None,
) -> PlotSpec:
    plot = _require_object(
        value,
        f"table {table_name!r}.plot",
        {"kind", "x", "y", "y_error", "title", "x_label", "y_label"},
    )
    required = {"kind", "x", "y", "y_error", "title", "x_label", "y_label"}
    missing = required - set(plot)
    if missing:
        raise DataImportError(f"table {table_name!r}.plot is missing: {', '.join(sorted(missing))}")
    kind = plot["kind"]
    if not isinstance(kind, str) or kind not in _PLOT_KINDS:
        raise DataImportError(f"Unsupported plot kind for table {table_name!r}: {kind!r}")
    raw_y = plot["y"]
    if not isinstance(raw_y, list):
        raise DataImportError(f"table {table_name!r}.plot.y must be an array")
    if plot["x"] is None:
        x_index = None
    else:
        x_index = _selector(plot["x"], table.columns, f"table {table_name!r}.plot.x")
    y_indices = tuple(_selector(selector, table.columns, f"table {table_name!r}.plot.y") for selector in raw_y)
    if len(set(y_indices)) != len(y_indices):
        raise DataImportError(f"table {table_name!r}.plot.y contains duplicate columns")
    if kind == "none":
        if x_index is not None or y_indices or plot["y_error"]:
            raise DataImportError(f"A 'none' plot must have x=null, y=[], and y_error={{}} in table {table_name!r}")
    elif not y_indices:
        raise DataImportError(f"Plot {kind!r} requires at least one Y column in table {table_name!r}")
    if kind != "none" and x_index is not None and x_index in y_indices:
        raise DataImportError(f"The X column cannot also be a Y column in table {table_name!r}")
    for index in y_indices:
        if table.columns[index].kind != "number":
            raise DataImportError(f"Y column {table.columns[index].name!r} must be numeric")
    unit_columns = label_columns if label_columns is not None else table.columns
    selected_units = {unit_columns[index].unit.strip() for index in y_indices}
    if warnings is not None and len(selected_units) > 1:
        if "" in selected_units:
            warnings.append("Some selected Y columns lack units; shared Y-axis comparability is unverified.")
        else:
            warnings.append("Selected Y columns use mixed units on one axis; their values are not normalized.")
    if kind in {"line", "scatter", "line_symbol"} and x_index is not None:
        if table.columns[x_index].kind not in {"number", "datetime"}:
            raise DataImportError(
                f"Plot {kind!r} requires a numeric or datetime X column; "
                f"use 'column' for categorical X {table.columns[x_index].name!r}"
            )
    if kind == "column" and x_index is not None:
        if table.columns[x_index].kind not in {"number", "text", "datetime"}:
            raise DataImportError(f"Unsupported X column kind for column plot: {table.columns[x_index].kind}")

    raw_errors = plot["y_error"]
    if not isinstance(raw_errors, Mapping):
        raise DataImportError(f"table {table_name!r}.plot.y_error must be an object")
    error_indices: list[tuple[int, int]] = []
    all_error_columns: set[int] = set()
    y_by_original_name = {table.columns[index].name: index for index in y_indices}
    for y_name, error_selector in raw_errors.items():
        if not isinstance(y_name, str):
            raise DataImportError("plot.y_error keys must be original, unique source column names")
        if y_name not in y_by_original_name:
            raise DataImportError(f"plot.y_error key {y_name!r} must name one of the selected Y columns")
        y_index = y_by_original_name[y_name]
        error_index = _selector(error_selector, table.columns, f"plot.y_error[{y_name!r}]")
        if error_index == y_index or error_index == x_index or error_index in y_indices:
            raise DataImportError(f"Error column {table.columns[error_index].name!r} conflicts with X or Y")
        if error_index in all_error_columns:
            raise DataImportError("Each Y uncertainty must use a distinct error column")
        if table.columns[error_index].kind != "number":
            raise DataImportError(f"Error column {table.columns[error_index].name!r} must be numeric")
        y_unit = unit_columns[y_index].unit.strip()
        error_unit = unit_columns[error_index].unit.strip()
        if y_unit and error_unit and y_unit != error_unit:
            raise DataImportError(
                f"Y error unit {error_unit!r} does not match Y unit {y_unit!r} for "
                f"{table.columns[y_index].name!r}; no unit conversion is applied"
            )
        if bool(y_unit) != bool(error_unit) and warnings is not None:
            warnings.append(
                f"Cannot verify that Y error column {table.columns[error_index].name!r} and "
                f"Y column {table.columns[y_index].name!r} use matching units; no unit conversion is applied."
            )
        for row_index, raw_value in enumerate(table.columns[error_index].values):
            if raw_value is None:
                continue
            uncertainty = _numeric_decimal(
                raw_value,
                context=f"{table_name}.{table.columns[error_index].name}[row {row_index}]",
            )
            if uncertainty < 0:
                raise DataImportError(
                    f"Error column {table.columns[error_index].name!r} contains a negative value at row {row_index}"
                )
        all_error_columns.add(error_index)
        error_indices.append((y_index, error_index))

    title = _require_string(plot["title"], f"table {table_name!r}.plot.title", nonempty=False)
    x_label = _require_string(plot["x_label"], f"table {table_name!r}.plot.x_label", nonempty=False)
    y_label = _require_string(plot["y_label"], f"table {table_name!r}.plot.y_label", nonempty=False)
    return PlotSpec(kind, x_index, y_indices, tuple(error_indices), title, x_label, y_label)


def _label_columns(raw: Any, table: DataTable, table_name: str) -> tuple[DataColumn, ...]:
    if not isinstance(raw, Mapping):
        raise DataImportError(f"table {table_name!r}.column_labels must be an object")
    source_names = {column.name for column in table.columns}
    unknown_names = set(raw) - source_names
    if unknown_names:
        raise DataImportError(
            f"Unknown column_labels source column(s) in {table_name!r}: {', '.join(sorted(map(str, unknown_names)))}"
        )
    output_columns: list[DataColumn] = []
    names: set[str] = set()
    for column in table.columns:
        entry = raw.get(column.name, {})
        if not isinstance(entry, Mapping):
            raise DataImportError(f"column_labels[{column.name!r}] must be an object")
        _unknown_keys(entry, {"name", "unit"}, f"column_labels[{column.name!r}]")
        name = entry.get("name", column.name)
        unit = entry.get("unit", column.unit)
        name = _require_string(name, f"column_labels[{column.name!r}].name")
        unit = _require_string(unit, f"column_labels[{column.name!r}].unit", nonempty=False)
        if name.casefold() in names:
            raise DataImportError(f"Output column labels must be unique in table {table_name!r}: {name!r}")
        names.add(name.casefold())
        output_columns.append(
            replace(column, name=name, unit=unit, original_name=column.original_name or column.name)
        )
    return tuple(output_columns)


def _canonical_output(value: Any, base_dir: Path | None) -> tuple[Path, str, bool, bool]:
    output = _require_object(value, "output", {"path", "format", "overwrite", "keep_open"})
    required = {"path", "format", "overwrite", "keep_open"}
    missing = required - set(output)
    if missing:
        raise DataImportError(f"output is missing: {', '.join(sorted(missing))}")
    raw_path = _require_string(output["path"], "output.path")
    file_format = output["format"]
    if not isinstance(file_format, str) or file_format not in _FORMAT_SUFFIX:
        raise DataImportError("output.format must be 'opju' or 'xlsx'")
    overwrite = output["overwrite"]
    keep_open = output["keep_open"]
    if not isinstance(overwrite, bool) or not isinstance(keep_open, bool):
        raise DataImportError("output.overwrite and output.keep_open must be booleans")
    if file_format == "xlsx" and keep_open:
        raise DataImportError("output.keep_open is only valid for opju output")
    target = _absolute(raw_path, base_dir)
    if target.suffix.casefold() != _FORMAT_SUFFIX[file_format]:
        raise DataImportError(f"output.path suffix must be {_FORMAT_SUFFIX[file_format]} for format {file_format!r}")
    if target.exists() and target.is_dir():
        raise DataImportError(f"Output path is a directory: {target}")
    ancestor = target.parent
    while not ancestor.exists() and ancestor != ancestor.parent:
        ancestor = ancestor.parent
    if ancestor.exists() and not ancestor.is_dir():
        raise DataImportError(f"Output parent is not a directory: {ancestor}")
    return target, file_format, overwrite, keep_open


def _source_request(source: Any, base_dir: Path | None, table_index: int) -> tuple[Path, str, dict[str, Any]]:
    item = _require_object(source, f"tables[{table_index}].source", {"path", "sha256", "options"})
    required = {"path", "sha256", "options"}
    missing = required - set(item)
    if missing:
        raise DataImportError(f"tables[{table_index}].source is missing: {', '.join(sorted(missing))}")
    raw_path = _require_string(item["path"], f"tables[{table_index}].source.path")
    source_path = _absolute(raw_path, base_dir)
    if not source_path.is_file():
        raise DataImportError(f"Source file does not exist or is not a file: {source_path}")
    expected_hash = item["sha256"]
    if not isinstance(expected_hash, str) or not _SHA256_RE.fullmatch(expected_hash.lower()):
        raise DataImportError(f"tables[{table_index}].source.sha256 must be a 64-character SHA-256")
    options = _validate_options(item["options"], replayable=True)
    return source_path, expected_hash.lower(), options


def _load_and_verify_source(
    source: Any, base_dir: Path | None, table_index: int, cache: SourceCache | None = None,
    preloaded: DataTable | None = None,
) -> DataTable:
    source_path, expected_hash, options = _source_request(source, base_dir, table_index)
    loaded = [preloaded] if preloaded is not None else (
        cache.read if cache is not None else _read_tables
    )(source_path, **options)
    if not isinstance(loaded, list) or len(loaded) != 1:
        count = len(loaded) if isinstance(loaded, list) else "invalid"
        raise DataImportError(
            f"tables[{table_index}] must resolve to exactly one source table; reader returned {count}. "
            "Select a source sheet explicitly in source.options.sheet."
        )
    table = loaded[0]
    _validate_table_shape(table)
    actual_path, actual_hash = _source_identity(table, source_path)
    if actual_path != source_path:
        raise DataImportError(f"Reader returned a different source file for tables[{table_index}]")
    if actual_hash != expected_hash.lower():
        raise DataImportError(f"Source changed since inspection: {source_path}")
    actual_options = _read_options_for_table(table)
    if actual_options != options:
        raise DataImportError(f"Reader options are not replayable for {source_path}; inspect the source again")
    # Multiple plan items may use the same preloaded sheet with different plots.
    # Keep its immutable values shared, but isolate each item's provenance options.
    return replace(table, read_options=deepcopy(table.read_options)) if preloaded is not None else table


def _workbook_group_key(path: Path, digest: str, options: dict[str, Any]) -> tuple[Path, str, str]:
    return path, digest, json.dumps(dict(options, sheet=None), sort_keys=True)


def _workbook_source_groups(raw_tables: list[Any], base_dir: Path | None):
    """Group metadata only; parse each workbook when its first table is needed."""
    groups: dict[tuple[Path, str, str], tuple[dict[str, Any], list[str]]] = {}
    for index, raw in enumerate(raw_tables):
        if not isinstance(raw, Mapping) or "source" not in raw:
            continue  # The ordinary validator supplies the table-specific error.
        path, digest, options = _source_request(raw["source"], base_dir, index)
        if path.suffix.casefold() not in {".xlsx", ".xlsm", ".xls"} or options["sheet"] is None:
            continue
        common = dict(options, sheet=None)
        key = _workbook_group_key(path, digest, common)
        _, sheets = groups.setdefault(key, (common, []))
        if options["sheet"] not in sheets:
            sheets.append(options["sheet"])
    return {key: value for key, value in groups.items() if len(value[1]) > 1}


def prepare_plan(
    plan: dict[str, Any], base_dir: Path | None = None, *, cache: SourceCache | None = None
) -> PreparedImport:
    """Validate a declarative plan and resolve every source without writing or starting Origin."""
    with cache.operation() if cache is not None else nullcontext():
        return _prepare_plan(plan, base_dir, cache)


def _prepare_plan(plan: dict[str, Any], base_dir: Path | None, cache: SourceCache | None) -> PreparedImport:
    root = _require_object(plan, "plan", {"schema_version", "output", "tables"})
    required = {"schema_version", "output", "tables"}
    missing = required - set(root)
    if missing:
        raise DataImportError(f"plan is missing: {', '.join(sorted(missing))}")
    if type(root["schema_version"]) is not int or root["schema_version"] != SCHEMA_VERSION:
        raise DataImportError(f"Unsupported import plan schema_version: {root['schema_version']!r}")
    if base_dir is not None and not isinstance(base_dir, Path):
        raise DataImportError("base_dir must be a pathlib.Path or None")
    resolved_base = _absolute(base_dir) if base_dir is not None else None
    output_path, file_format, overwrite, keep_open = _canonical_output(root["output"], resolved_base)

    raw_tables = root["tables"]
    if not isinstance(raw_tables, list) or not raw_tables:
        raise DataImportError("plan.tables must be a non-empty array")
    groups = _workbook_source_groups(raw_tables, resolved_base) if cache is not None else {}
    # These references share immutable values with the prepared result, including
    # oversized workbooks. They disappear with this call, independent of the LRU.
    loaded_groups: dict[tuple[Path, str, str], dict[str | None, DataTable]] = {}
    planned: list[PlannedTable] = []
    names: set[str] = set()
    input_paths: set[Path] = set()
    path_sheet_counts: dict[Path, int] = {}
    for index, raw in enumerate(raw_tables):
        table_item = _require_object(
            raw,
            f"tables[{index}]",
            {"source", "name", "plot", "column_labels"},
        )
        missing_table = {"source", "name", "plot", "column_labels"} - set(table_item)
        if missing_table:
            raise DataImportError(f"tables[{index}] is missing: {', '.join(sorted(missing_table))}")
        name = _require_string(table_item["name"], f"tables[{index}].name")
        name_key = name.strip().casefold()
        if name_key in names:
            raise DataImportError(f"Target table names must be unique: {name!r}")
        names.add(name_key)
        source_mapping = _require_object(
            table_item["source"], f"tables[{index}].source", {"path", "sha256", "options"}
        )
        raw_source_path = _require_string(source_mapping.get("path"), f"tables[{index}].source.path")
        source_path = _absolute(raw_source_path, resolved_base)
        input_paths.add(source_path)
        path_sheet_counts[source_path] = path_sheet_counts.get(source_path, 0) + 1
        preloaded = None
        if cache is not None:
            _path, digest, options = _source_request(source_mapping, resolved_base, index)
            key = _workbook_group_key(source_path, digest, options)
            if key in groups and key not in loaded_groups:
                common, sheets = groups[key]
                loaded_groups[key] = {
                    item.source_sheet: item
                    for item in cache.read_sheets(source_path, tuple(sheets), common)
                }
            preloaded = loaded_groups.get(key, {}).get(options["sheet"])
        table = _load_and_verify_source(table_item["source"], resolved_base, index, cache, preloaded)
        columns = _label_columns(table_item["column_labels"], table, name)
        plot_warnings: list[str] = []
        plot = _validate_plot(table_item["plot"], table, name, plot_warnings, columns)
        prepared_table = replace(table, columns=columns, warnings=tuple(table.warnings) + tuple(plot_warnings))
        planned.append(PlannedTable(prepared_table, name.strip(), plot))

    conflicts = sorted(str(path) for path in input_paths if path == output_path)
    if conflicts:
        raise DataImportError(f"Output path conflicts with an input source: {conflicts[0]}")
    if output_path.exists() and not overwrite:
        raise DataImportError(f"Output already exists and overwrite is false: {output_path}")
    for source_path, count in path_sheet_counts.items():
        if count > 1 and source_path.suffix.casefold() in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
            for item in raw_tables:
                source = item["source"]
                if _absolute(source["path"], resolved_base) == source_path and source["options"].get("sheet") is None:
                    raise DataImportError(
                        f"source.options.sheet must identify each worksheet when a plan includes multiple tables from {source_path}"
                    )

    warnings = tuple(warning for planned_table in planned for warning in planned_table.table.warnings)
    return PreparedImport(tuple(planned), output_path, file_format, overwrite, keep_open, warnings)


def describe_prepared(prepared: PreparedImport) -> dict[str, Any]:
    """Return the validated import in a JSON-safe shape shared by CLI and GUI."""
    if not isinstance(prepared, PreparedImport):
        raise DataImportError("prepared must be a PreparedImport")
    from .exporter import validate_prepared_import
    conversion_warnings = validate_prepared_import(prepared)
    tables = []
    graphs = []
    for item in prepared.tables:
        table = item.table
        plot = item.plot
        source = {
            "path": str(_absolute(table.source)),
            "sheet": table.source_sheet,
            "sha256": table.source_hash.lower(),
        }
        columns = [
            {
                "index": index,
                "source_name": column.original_name or column.name,
                "name": column.name,
                "kind": column.kind,
                "unit": column.unit,
            }
            for index, column in enumerate(table.columns)
        ]
        plot_summary = {
            "kind": plot.kind,
            "x": plot.x,
            "y": list(plot.y),
            "y_error": {table.columns[y].name: table.columns[error].name for y, error in plot.y_error},
            "title": plot.title,
            "x_label": plot.x_label,
            "y_label": plot.y_label,
        }
        tables.append(
            {
                "source": source,
                "name": item.name,
                "n_rows": table.n_rows,
                "columns": columns,
                "plot": plot_summary,
            }
        )
        if plot.kind != "none":
            graphs.append({"table": item.name, **plot_summary})
    return {
        "schema_version": SCHEMA_VERSION,
        "valid": True,
        "output": {
            "path": str(prepared.output),
            "format": prepared.format,
            "overwrite": prepared.overwrite,
            "keep_open": prepared.keep_open,
        },
        "tables": tables,
        "graphs": graphs,
        "warnings": conversion_warnings,
    }
