"""Loss-aware readers for common scientific tabular files.

The reader layer deliberately stores cell values as text.  It infers a column
kind for downstream Origin conversion, while preserving the source spelling
(except Fortran ``D`` exponents, which are normalized to ``E``).
"""

from __future__ import annotations

import csv
import hashlib
import io
import json
import math
import re
from datetime import date, datetime, time
from decimal import Decimal
from pathlib import Path
from typing import Any

from .models import DataColumn, DataImportError, DataTable


SUPPORTED_EXTENSIONS = frozenset(
    {".txt", ".dat", ".xy", ".csv", ".tsv", ".xlsx", ".xlsm", ".xls", ".json", ".jsonl", ".ndjson"}
)
_TEXT_EXTENSIONS = frozenset({".txt", ".dat", ".xy", ".csv", ".tsv"})
_JSONL_EXTENSIONS = frozenset({".jsonl", ".ndjson"})
_DELIMITERS = {",": ",", ";": ";", "\t": "\t", "whitespace": "whitespace"}
_NUMERIC_RE = re.compile(
    r"^[+-]?(?:(?:\d+(?:\.\d*)?)|(?:\.\d+))(?:[eEdD][+-]?\d+)?$"
)
_LEADING_ZERO_RE = re.compile(r"^[+-]?0\d+$")
_UNIT_PARENS_RE = re.compile(r"^(.*?)\s*\(([^()]*)\)\s*$")
_UNIT_BRACKETS_RE = re.compile(r"^(.*?)\s*\[([^\[\]]*)\]\s*$")


class _JSONObject(dict[str, Any]):
    """A parsed JSON object that remembers keys Python's default parser loses."""

    def __init__(self, pairs: list[tuple[str, Any]]) -> None:
        super().__init__()
        self.duplicate_keys: list[str] = []
        for key, value in pairs:
            if key in self and key not in self.duplicate_keys:
                self.duplicate_keys.append(key)
            self[key] = value


def _natural_key(path: Path) -> tuple[Any, ...]:
    return tuple(
        int(part) if part.isdigit() else part.casefold()
        for part in re.split(r"(\d+)", path.as_posix())
    )


def discover_files(paths: list[Path]) -> list[Path]:
    """Expand files and folders recursively, returning supported files naturally sorted."""

    if not paths:
        raise DataImportError("至少需要提供一个文件或目录。")

    found: dict[Path, None] = {}
    for raw_path in paths:
        path = Path(raw_path).expanduser()
        if not path.exists():
            raise DataImportError(f"输入路径不存在：{path}")
        resolved = path.resolve()
        if resolved.is_file():
            if resolved.suffix.casefold() not in SUPPORTED_EXTENSIONS:
                raise DataImportError(f"不支持的文件类型：{resolved.name}")
            found[resolved] = None
            continue
        if not resolved.is_dir():
            raise DataImportError(f"输入路径不是文件或目录：{resolved}")

        candidates = [
            child.resolve()
            for child in resolved.rglob("*")
            if child.is_file() and child.suffix.casefold() in SUPPORTED_EXTENSIONS
        ]
        if not candidates:
            raise DataImportError(f"目录中没有可导入的数据文件：{resolved}")
        for candidate in candidates:
            found[candidate] = None

    return sorted(found, key=_natural_key)


def read_tables(path: Path, **options: Any) -> list[DataTable]:
    """Read a supported source file into one or more rectangular tables.

    Supported options are ``sheet``, ``header``, ``skip_rows``, ``delimiter``,
    ``encoding``, ``missing_values``, and ``formula_policy``.  The normalized
    settings are copied to each returned table so that a caller can replay the
    import decision later.
    """

    resolved = Path(path).expanduser()
    if not resolved.exists() or not resolved.is_file():
        raise DataImportError(f"数据文件不存在或不是文件：{resolved}")
    resolved = resolved.resolve()
    suffix = resolved.suffix.casefold()
    if suffix not in SUPPORTED_EXTENSIONS:
        raise DataImportError(f"不支持的文件类型：{resolved.name}")

    settings = _normalize_options(options)
    if settings["sheet"] is not None and suffix not in {".xlsx", ".xlsm", ".xls"}:
        raise DataImportError("sheet 选项仅适用于 Excel 文件。")
    try:
        source_bytes = resolved.read_bytes()
    except OSError as exc:
        raise DataImportError(f"无法读取文件 {resolved.name}：{exc}") from exc
    if not source_bytes:
        raise DataImportError(f"文件为空：{resolved.name}")
    source_hash = hashlib.sha256(source_bytes).hexdigest()

    if suffix in _TEXT_EXTENSIONS:
        text, actual_encoding = _decode_text(source_bytes, settings["encoding"], resolved)
        rows, line_numbers, actual_delimiter = _read_delimited_rows(
            text, settings["delimiter"], settings["skip_rows"], resolved
        )
        settings["encoding"] = actual_encoding
        settings["delimiter"] = actual_delimiter
        settings["sheet"] = None
        return [
            _build_table(
                resolved,
                None,
                resolved.stem,
                rows,
                line_numbers,
                source_hash,
                settings,
                format_kind="delimited",
                header_mode=settings["header"],
                missing_values=set(settings["missing_values"]),
            )
        ]

    if suffix in {".xlsx", ".xlsm", ".xls"}:
        return _read_excel(resolved, source_bytes, source_hash, settings)

    if suffix in _JSONL_EXTENSIONS:
        text, actual_encoding = _decode_text(source_bytes, settings["encoding"], resolved)
        rows, names, line_numbers = _read_jsonl(text, settings["skip_rows"], resolved)
        settings["encoding"] = actual_encoding
        settings["sheet"] = None
        return [
            _build_structured_table(
                resolved, None, resolved.stem, names, rows, line_numbers,
                source_hash, settings, set(settings["missing_values"]),
            )
        ]

    text, actual_encoding = _decode_text(source_bytes, settings["encoding"], resolved)
    rows, names, line_numbers = _read_json(text, settings["skip_rows"], resolved)
    settings["encoding"] = actual_encoding
    settings["sheet"] = None
    return [
        _build_structured_table(
            resolved, None, resolved.stem, names, rows, line_numbers,
            source_hash, settings, set(settings["missing_values"]),
        )
    ]


def _normalize_options(options: dict[str, Any]) -> dict[str, Any]:
    allowed = {"sheet", "header", "skip_rows", "delimiter", "encoding", "missing_values", "formula_policy"}
    unknown = sorted(set(options) - allowed)
    if unknown:
        raise DataImportError(f"未知的读取选项：{', '.join(unknown)}")

    sheet = options.get("sheet")
    if sheet is not None and (not isinstance(sheet, str) or not sheet):
        raise DataImportError("sheet 必须是非空工作表名或 None。")
    header = options.get("header", "auto")
    if header not in {"auto", True, False}:
        raise DataImportError("header 必须是 'auto'、True 或 False。")
    skip_rows = options.get("skip_rows", 0)
    if isinstance(skip_rows, bool) or not isinstance(skip_rows, int) or skip_rows < 0:
        raise DataImportError("skip_rows 必须是非负整数。")
    delimiter = options.get("delimiter", "auto")
    if delimiter != "auto" and delimiter not in _DELIMITERS:
        raise DataImportError("delimiter 必须是 'auto'、'whitespace'、','、';' 或 '\\t'。")
    encoding = options.get("encoding", "auto")
    if not isinstance(encoding, str) or not encoding:
        raise DataImportError("encoding 必须是 'auto' 或明确的编码名。")
    missing_values = options.get("missing_values", [""])
    if not isinstance(missing_values, (list, tuple)) or any(not isinstance(item, str) for item in missing_values):
        raise DataImportError("missing_values 必须是字符串列表。")
    formula_policy = options.get("formula_policy", "cached")
    if formula_policy not in {"cached", "text"}:
        raise DataImportError("formula_policy 必须是 'cached' 或 'text'。")
    return {
        "sheet": sheet,
        "header": header,
        "skip_rows": skip_rows,
        "delimiter": delimiter,
        "encoding": encoding,
        "missing_values": list(missing_values),
        "formula_policy": formula_policy,
    }


def _decode_text(data: bytes, encoding: str, path: Path) -> tuple[str, str]:
    if encoding != "auto":
        try:
            return data.decode(encoding), encoding
        except (LookupError, UnicodeDecodeError) as exc:
            raise DataImportError(f"{path.name} 无法按编码 {encoding} 解码：{exc}") from exc

    # BOM-specific codecs must precede UTF-8; GB18030 covers both GBK and GB2312.
    attempts = ("utf-8-sig", "utf-32", "utf-16", "gb18030", "cp1252")
    failures: list[str] = []
    for candidate in attempts:
        try:
            return data.decode(candidate), candidate
        except (LookupError, UnicodeDecodeError) as exc:
            failures.append(f"{candidate}: {exc}")
    raise DataImportError(f"{path.name} 无法识别文本编码（{'；'.join(failures)}）。")


def _detect_delimiter(text: str, requested: str, path: Path) -> str:
    if requested != "auto":
        return requested
    sample = "\n".join(line for line in text.splitlines()[:30] if line.strip())
    if not sample:
        raise DataImportError(f"文件没有可解析的数据行：{path.name}")
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
        return dialect.delimiter
    except csv.Error:
        if any(ch in sample for ch in ",;\t"):
            counts = {ch: sample.count(ch) for ch in ",;\t"}
            return max(counts, key=counts.get)
        return "whitespace"


def _read_delimited_rows(
    text: str, requested_delimiter: str, skip_rows: int, path: Path
) -> tuple[list[list[str]], list[int], str]:
    lines = text.splitlines(keepends=True)
    if skip_rows >= len(lines):
        raise DataImportError(f"{path.name} 跳过 {skip_rows} 行后没有数据。")
    remaining = "".join(lines[skip_rows:])
    delimiter = _detect_delimiter(remaining, requested_delimiter, path)
    rows: list[list[str]] = []
    line_numbers: list[int] = []
    if delimiter == "whitespace":
        for index, line in enumerate(remaining.splitlines(), start=skip_rows + 1):
            rows.append(line.split())
            line_numbers.append(index)
        return rows, line_numbers, delimiter

    reader = csv.reader(io.StringIO(remaining, newline=""), delimiter=delimiter, strict=True)
    try:
        for row in reader:
            rows.append(row)
            line_numbers.append(skip_rows + reader.line_num)
    except csv.Error as exc:
        raise DataImportError(f"{path.name} 第 {skip_rows + reader.line_num} 行分隔符格式无效：{exc}") from exc
    return rows, line_numbers, delimiter


def _read_json(text: str, skip_rows: int, path: Path) -> tuple[list[list[Any]], list[str], list[int]]:
    lines = text.splitlines()
    if skip_rows:
        lines = lines[skip_rows:]
    payload_text = "\n".join(lines)
    if not payload_text.strip():
        raise DataImportError(f"{path.name} 跳过 {skip_rows} 行后没有 JSON 数据。")
    try:
        payload = json.loads(
            payload_text,
            parse_float=Decimal,
            object_pairs_hook=_JSONObject,
        )
    except json.JSONDecodeError as exc:
        raise DataImportError(f"{path.name} 第 {skip_rows + exc.lineno} 行 JSON 格式无效：{exc.msg}") from exc
    if isinstance(payload, list):
        for record_number, record in enumerate(payload, start=1):
            _reject_duplicate_json_keys(record, path, f"第 {record_number} 条记录")
    else:
        _reject_duplicate_json_keys(payload, path, "JSON 对象")

    if isinstance(payload, dict) and "columns" in payload and "data" in payload:
        names = _json_column_names(payload["columns"], path)
        data = payload["data"]
        if not isinstance(data, list):
            raise DataImportError(f"{path.name} 的 data 必须是二维数组。")
        rows = []
        line_numbers = []
        for row_index, row in enumerate(data, start=1):
            if not isinstance(row, list) or len(row) != len(names):
                raise DataImportError(f"{path.name} 的 data 第 {row_index} 行不是 {len(names)} 列矩形数据。")
            rows.append([_scalar_json_value(value, path, row_index) for value in row])
            line_numbers.append(row_index)
        return rows, names, line_numbers

    if isinstance(payload, list):
        if not payload:
            raise DataImportError(f"{path.name} 的 JSON 数组为空。")
        if not all(isinstance(record, dict) for record in payload):
            raise DataImportError(f"{path.name} 的 JSON 数组必须由记录对象组成。")
        names = list(payload[0])
        if not names:
            raise DataImportError(f"{path.name} 的 JSON 记录没有字段。")
        rows = []
        for row_index, record in enumerate(payload, start=1):
            if set(record) != set(names):
                raise DataImportError(f"{path.name} 第 {row_index} 条记录字段不一致，数据不是矩形表。")
            rows.append([_scalar_json_value(record[name], path, row_index) for name in names])
        return rows, [str(name) for name in names], list(range(1, len(rows) + 1))

    if isinstance(payload, dict) and payload and all(isinstance(values, list) for values in payload.values()):
        names = [str(name) for name in payload]
        lengths = {len(values) for values in payload.values()}
        if len(lengths) != 1:
            raise DataImportError(f"{path.name} 的列数组长度不一致，数据不是矩形表。")
        row_count = lengths.pop()
        rows = [
            [_scalar_json_value(payload[name][row_index], path, row_index + 1) for name in payload]
            for row_index in range(row_count)
        ]
        return rows, names, list(range(1, row_count + 1))

    raise DataImportError(
        f"{path.name} JSON 顶层必须是记录数组、列数组对象，或包含 columns 与 data 的对象。"
    )


def _json_column_names(columns: Any, path: Path) -> list[str]:
    if not isinstance(columns, list) or not columns:
        raise DataImportError(f"{path.name} 的 columns 必须是非空名称数组。")
    names: list[str] = []
    for index, column in enumerate(columns, start=1):
        if isinstance(column, str):
            names.append(column)
        elif isinstance(column, dict) and isinstance(column.get("name"), str):
            name = column["name"]
            unit = column.get("unit")
            if isinstance(unit, str) and unit:
                name = f"{name} ({unit})"
            names.append(name)
        else:
            raise DataImportError(f"{path.name} 的 columns 第 {index} 项不是字符串或带 name 的对象。")
    return names


def _read_jsonl(text: str, skip_rows: int, path: Path) -> tuple[list[list[Any]], list[str], list[int]]:
    lines = text.splitlines()
    if skip_rows >= len(lines):
        raise DataImportError(f"{path.name} 跳过 {skip_rows} 行后没有 JSONL 数据。")
    records: list[dict[str, Any]] = []
    line_numbers: list[int] = []
    names: list[str] | None = None
    for physical_line, line in enumerate(lines[skip_rows:], start=skip_rows + 1):
        if not line.strip():
            raise DataImportError(f"{path.name} 第 {physical_line} 行为空；JSONL 每行都必须是一条记录。")
        try:
            record = json.loads(
                line,
                parse_float=Decimal,
                object_pairs_hook=_JSONObject,
            )
        except json.JSONDecodeError as exc:
            raise DataImportError(f"{path.name} 第 {physical_line} 行 JSONL 格式无效：{exc.msg}") from exc
        _reject_duplicate_json_keys(record, path, f"第 {physical_line} 行")
        if not isinstance(record, dict) or not record:
            raise DataImportError(f"{path.name} 第 {physical_line} 行必须是非空 JSON 对象记录。")
        if names is None:
            names = list(record)
        elif set(record) != set(names):
            raise DataImportError(f"{path.name} 第 {physical_line} 行记录字段不一致，数据不是矩形表。")
        records.append(record)
        line_numbers.append(physical_line)
    if not records or names is None:
        raise DataImportError(f"{path.name} 没有 JSONL 记录。")
    rows = [
        [_scalar_json_value(record[name], path, line_numbers[row_index]) for name in names]
        for row_index, record in enumerate(records)
    ]
    return rows, [str(name) for name in names], line_numbers


def _reject_duplicate_json_keys(value: Any, path: Path, context: str) -> None:
    if isinstance(value, _JSONObject):
        if value.duplicate_keys:
            keys = ", ".join(repr(key) for key in value.duplicate_keys)
            raise DataImportError(f"{path.name} {context}含有重复 JSON 键 {keys}；拒绝静默覆盖数据。")
        for nested in value.values():
            _reject_duplicate_json_keys(nested, path, context)
    elif isinstance(value, list):
        for nested in value:
            _reject_duplicate_json_keys(nested, path, context)


def _scalar_json_value(value: Any, path: Path, row_number: int) -> Any:
    if isinstance(value, (dict, list)):
        raise DataImportError(f"{path.name} 第 {row_number} 行含有嵌套 JSON 值，无法无损转换为单元格。")
    if isinstance(value, bool):
        return "true" if value else "false"
    if value is None:
        return None
    if isinstance(value, Decimal):
        return str(value)
    return str(value)


def _read_excel(path: Path, source_bytes: bytes, source_hash: str, settings: dict[str, Any]) -> list[DataTable]:
    try:
        import openpyxl
    except ImportError as exc:
        raise DataImportError("读取 xlsx/xlsm 需要安装 openpyxl。") from exc

    if path.suffix.casefold() == ".xls":
        try:
            import xlrd
        except ImportError as exc:
            raise DataImportError("读取旧版 .xls 文件需要安装 xlrd。") from exc
        return _read_xls(path, source_bytes, source_hash, settings, xlrd)

    formula_policy = settings["formula_policy"]
    try:
        workbook = openpyxl.load_workbook(
            io.BytesIO(source_bytes), read_only=True, data_only=(formula_policy == "cached")
        )
        formula_workbook = (
            openpyxl.load_workbook(io.BytesIO(source_bytes), read_only=True, data_only=False)
            if formula_policy == "cached"
            else None
        )
    except Exception as exc:
        raise DataImportError(f"无法读取 Excel 文件 {path.name}：{exc}") from exc
    try:
        if settings["sheet"] is not None:
            if settings["sheet"] not in workbook.sheetnames:
                raise DataImportError(f"{path.name} 中不存在工作表：{settings['sheet']}")
            sheet_names = [settings["sheet"]]
        else:
            sheet_names = list(workbook.sheetnames)

        tables: list[DataTable] = []
        for sheet_name in sheet_names:
            worksheet = workbook[sheet_name]
            formula_sheet = formula_workbook[sheet_name] if formula_workbook is not None else None
            rows, line_numbers = _excel_rows(
                worksheet, formula_sheet, settings["skip_rows"], formula_policy, path
            )
            if not rows or not any(any(value is not None for value in row) for row in rows):
                continue
            table_settings = dict(settings)
            table_settings["sheet"] = sheet_name
            table = _build_table(
                path,
                sheet_name,
                f"{path.stem}_{sheet_name}",
                rows,
                line_numbers,
                source_hash,
                table_settings,
                format_kind="excel",
                header_mode=settings["header"],
                missing_values=set(settings["missing_values"]),
            )
            tables.append(table)
        if not tables:
            raise DataImportError(f"{path.name} 中没有非空工作表。")
        return tables
    finally:
        workbook.close()
        if formula_workbook is not None:
            formula_workbook.close()


def _excel_rows(
    worksheet: Any,
    formula_worksheet: Any | None,
    skip_rows: int,
    formula_policy: str,
    path: Path,
) -> tuple[list[list[Any]], list[int]]:
    rows: list[list[Any]] = []
    line_numbers: list[int] = []
    formula_rows = (
        iter(formula_worksheet.iter_rows(min_row=skip_rows + 1))
        if formula_worksheet is not None
        else None
    )
    for row_number, cells in enumerate(worksheet.iter_rows(min_row=skip_rows + 1), start=skip_rows + 1):
        values: list[Any] = []
        formula_cells = next(formula_rows, None) if formula_rows is not None else None
        if formula_rows is not None and formula_cells is None:
            raise DataImportError(
                f"{path.name} 工作表 {worksheet.title} 的公式视图行数与缓存视图不一致。"
            )
        for column_index, cell in enumerate(cells, start=1):
            formula_cell = formula_cells[column_index - 1] if formula_cells and column_index <= len(formula_cells) else None
            if formula_policy == "cached" and formula_cell is not None and formula_cell.data_type == "f" and cell.value is None:
                coordinate = formula_cell.coordinate
                raise DataImportError(
                    f"{path.name} 工作表 {worksheet.title} 单元格 {coordinate} 的公式没有缓存结果；"
                    "可用 formula_policy='text' 保留公式文本。"
                )
            value = cell.value
            if formula_policy == "text" and cell.data_type == "f":
                values.append(str(value))
            elif isinstance(value, (datetime, date, time)):
                values.append(value.isoformat())
            elif isinstance(value, bool):
                values.append("true" if value else "false")
            elif isinstance(value, (int, float)):
                if isinstance(value, float) and not math.isfinite(value):
                    raise DataImportError(
                        f"{path.name} 工作表 {worksheet.title} 第 {row_number} 行第 {column_index} 列含非有限数值。"
                    )
                zero_format = re.fullmatch(r"0+", cell.number_format or "")
                if zero_format and float(value).is_integer():
                    values.append(str(int(value)).zfill(len(cell.number_format)))
                else:
                    values.append(str(value))
            elif value is None:
                values.append(None)
            else:
                values.append(str(value))
        rows.append(values)
        line_numbers.append(row_number)

    # Excel often reports styled trailing rows as part of its used range; they are not records.
    while rows and not any(value is not None for value in rows[-1]):
        rows.pop()
        line_numbers.pop()
    while rows and not any(value is not None for value in rows[0]):
        rows.pop(0)
        line_numbers.pop(0)
    return rows, line_numbers


def _read_xls(
    path: Path, source_bytes: bytes, source_hash: str, settings: dict[str, Any], xlrd: Any
) -> list[DataTable]:
    if settings["formula_policy"] == "text":
        raise DataImportError("xlrd 无法读取 .xls 公式文本；formula_policy='text' 仅适用于 .xlsx/.xlsm。")
    try:
        workbook = xlrd.open_workbook(file_contents=source_bytes, formatting_info=True)
    except Exception as exc:
        raise DataImportError(f"无法读取 XLS 文件 {path.name}：{exc}") from exc
    if settings["sheet"] is not None:
        try:
            sheet = workbook.sheet_by_name(settings["sheet"])
        except xlrd.XLRDError as exc:
            raise DataImportError(f"{path.name} 中不存在工作表：{settings['sheet']}") from exc
        sheets = [sheet]
    else:
        sheets = [workbook.sheet_by_index(index) for index in range(workbook.nsheets)]

    tables: list[DataTable] = []
    for sheet in sheets:
        rows: list[list[Any]] = []
        line_numbers: list[int] = []
        for row_index in range(settings["skip_rows"], sheet.nrows):
            row: list[Any] = []
            for column_index in range(sheet.ncols):
                cell = sheet.cell(row_index, column_index)
                value = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    date_value = xlrd.xldate_as_datetime(value, workbook.datemode)
                    row.append(date_value.isoformat())
                elif cell.ctype in {xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK}:
                    row.append(None)
                elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                    row.append("true" if value else "false")
                elif cell.ctype == xlrd.XL_CELL_ERROR:
                    raise DataImportError(
                        f"{path.name} 工作表 {sheet.name} 第 {row_index + 1} 行第 {column_index + 1} 列包含 Excel 错误值。"
                    )
                elif cell.ctype == xlrd.XL_CELL_NUMBER:
                    if not math.isfinite(value):
                        raise DataImportError(
                            f"{path.name} 工作表 {sheet.name} 第 {row_index + 1} 行第 {column_index + 1} 列含非有限数值。"
                        )
                    xf = workbook.xf_list[sheet.cell_xf_index(row_index, column_index)]
                    number_format = workbook.format_map.get(xf.format_key)
                    format_string = number_format.format_str if number_format is not None else ""
                    if re.fullmatch(r"0+", format_string) and float(value).is_integer():
                        row.append(str(int(value)).zfill(len(format_string)))
                    else:
                        row.append(str(int(value)) if float(value).is_integer() else str(value))
                else:
                    row.append(str(value))
            rows.append(row)
            line_numbers.append(row_index + 1)
        while rows and not any(value is not None for value in rows[-1]):
            rows.pop()
            line_numbers.pop()
        while rows and not any(value is not None for value in rows[0]):
            rows.pop(0)
            line_numbers.pop(0)
        if rows:
            table_settings = dict(settings)
            table_settings["sheet"] = sheet.name
            tables.append(
                _build_table(
                    path, sheet.name, f"{path.stem}_{sheet.name}", rows, line_numbers,
                    source_hash, table_settings, format_kind="excel",
                    header_mode=settings["header"], missing_values=set(settings["missing_values"]),
                )
            )
    if not tables:
        raise DataImportError(f"{path.name} 中没有非空工作表。")
    return tables


def _build_structured_table(
    path: Path,
    source_sheet: str | None,
    name: str,
    names: list[str],
    rows: list[list[Any]],
    line_numbers: list[int],
    source_hash: str,
    settings: dict[str, Any],
    missing_values: set[str],
) -> DataTable:
    if settings["header"] is False:
        raise DataImportError(f"{path.name} 的 JSON 字段名是数据结构的一部分，不能设置 header=False。")
    settings["header"] = True
    if len(rows) != len(line_numbers):
        raise DataImportError(f"{path.name} 的内部行数不一致。")
    columns, warnings = _make_columns(names, rows, line_numbers, missing_values, path)
    return DataTable(path, source_sheet, name, tuple(columns), source_hash, dict(settings), tuple(warnings))


def _build_table(
    path: Path,
    source_sheet: str | None,
    name: str,
    rows: list[list[Any]],
    line_numbers: list[int],
    source_hash: str,
    settings: dict[str, Any],
    *,
    format_kind: str,
    header_mode: str | bool,
    missing_values: set[str],
) -> DataTable:
    if not rows:
        raise DataImportError(f"{path.name} 没有可用的数据行。")
    width = len(rows[0])
    if width == 0:
        raise DataImportError(f"{path.name} 第 {line_numbers[0]} 行为空记录，无法确定列数。")
    for row_index, row in enumerate(rows):
        if len(row) != width:
            physical_line = line_numbers[row_index] if row_index < len(line_numbers) else row_index + 1
            raise DataImportError(
                f"{path.name} 第 {physical_line} 行有 {len(row)} 列，预期 {width} 列；数据必须是矩形表。"
            )

    warnings: list[str] = []
    has_header = _resolve_header(rows, line_numbers, header_mode, path, warnings)
    if has_header:
        raw_names = ["" if value is None else str(value).strip() for value in rows[0]]
        data_rows = rows[1:]
        data_lines = line_numbers[1:]
        if not data_rows:
            raise DataImportError(f"{path.name} 只有列名，没有数据记录。")
    else:
        raw_names = [f"Column {index + 1}" for index in range(width)]
        data_rows = rows
        data_lines = line_numbers

    columns, column_warnings = _make_columns(raw_names, data_rows, data_lines, missing_values, path)
    warnings.extend(column_warnings)
    settings["header"] = has_header
    return DataTable(path, source_sheet, name, tuple(columns), source_hash, dict(settings), tuple(warnings))


def _resolve_header(
    rows: list[list[Any]], line_numbers: list[int], mode: str | bool, path: Path, warnings: list[str]
) -> bool:
    if mode is True or mode is False:
        return mode
    first = rows[0]
    next_row = next((row for row in rows[1:] if any(value is not None and str(value).strip() for value in row)), None)
    present = [value for value in first if value is not None and str(value).strip()]
    if next_row is not None and present:
        first_is_text = all(not _is_numeric_token(str(value)) and not _is_iso_datetime(str(value)) for value in present)
        later_has_scalar = any(
            value is not None
            and (_is_numeric_token(str(value)) or _is_iso_datetime(str(value)))
            for value in next_row
        )
        later_row_is_data = all(
            value is None
            or not str(value).strip()
            or _is_numeric_token(str(value))
            or _is_iso_datetime(str(value))
            or _LEADING_ZERO_RE.fullmatch(str(value).strip()) is not None
            for value in next_row
        )
        header_has_units = any(
            _split_unit(str(value).strip())[1]
            for value in present
        )
        likely_header = first_is_text and later_has_scalar and (later_row_is_data or header_has_units)
        if likely_header:
            return True
    all_strings = all(
        value is None or not _is_numeric_token(str(value)) and not _is_iso_datetime(str(value))
        for row in rows
        for value in row
    )
    if all_strings:
        warnings.append("自动表头判断有歧义，未将首行当作表头；如需表头请设置 header=True。")
    elif present and all(not _is_numeric_token(str(value)) for value in present):
        warnings.append("首行不满足明确的表头特征，已保留为数据；如需表头请设置 header=True。")
    return False


def _is_numeric_token(value: str) -> bool:
    token = value.strip()
    if not token or _LEADING_ZERO_RE.fullmatch(token):
        return False
    if not _NUMERIC_RE.fullmatch(token):
        lowered = token.casefold()
        return lowered in {"nan", "+nan", "-nan", "inf", "+inf", "-inf", "infinity", "+infinity", "-infinity"}
    try:
        number = float(token.replace("D", "E").replace("d", "e"))
    except ValueError:
        return False
    return math.isfinite(number)


def _is_iso_datetime(value: str) -> bool:
    token = value.strip()
    if not token or _is_numeric_token(token):
        return False
    try:
        datetime.fromisoformat(token.replace("Z", "+00:00"))
        return True
    except ValueError:
        try:
            date.fromisoformat(token)
            return True
        except ValueError:
            return False


def _make_columns(
    raw_names: list[str],
    rows: list[list[Any]],
    line_numbers: list[int],
    missing_values: set[str],
    path: Path,
) -> tuple[list[DataColumn], list[str]]:
    if not raw_names:
        raise DataImportError(f"{path.name} 没有列。")
    if any(len(row) != len(raw_names) for row in rows):
        raise DataImportError(f"{path.name} 的列数和数据宽度不一致。")
    warnings: list[str] = []
    columns: list[DataColumn] = []
    used_names: set[str] = set()
    for column_index, raw_name in enumerate(raw_names):
        original_name = str(raw_name)
        base_name, unit = _split_unit(original_name.strip())
        if not base_name:
            base_name = f"Column {column_index + 1}"
        column_name = _unique_name(base_name, used_names)
        tokens: list[str | None] = []
        for row_index, row in enumerate(rows):
            raw_value = row[column_index]
            if raw_value is None:
                tokens.append(None)
                continue
            token = str(raw_value)
            normalized = token.strip()
            if normalized in missing_values:
                tokens.append(None)
                continue
            if _is_nonfinite_token(normalized):
                line = line_numbers[row_index] if row_index < len(line_numbers) else row_index + 1
                raise DataImportError(
                    f"{path.name} 第 {line} 行第 {column_index + 1} 列包含非有限数值 {normalized!r}。"
                )
            tokens.append(token)

        present = [token.strip() for token in tokens if token is not None]
        numeric = bool(present) and all(
            _is_numeric_token(token) and not _LEADING_ZERO_RE.fullmatch(token) for token in present
        )
        datetimes = bool(present) and all(_is_iso_datetime(token) for token in present)
        if numeric:
            kind = "number"
        elif datetimes:
            kind = "datetime"
        else:
            kind = "text"
            has_numeric = any(_is_numeric_token(token) for token in present)
            has_text = any(not _is_numeric_token(token) for token in present)
            if has_numeric and has_text:
                warnings.append(
                    f"列 {column_name} 同时包含数值和文本；所有单元格均按文本保留，缺失标记仅使用配置值。"
                )
        normalized_values: list[str | None] = []
        for token in tokens:
            if token is None:
                normalized_values.append(None)
            elif kind == "number":
                numeric_token = token.strip()
                numeric_token = re.sub(r"[dD](?=[+-]?\d+$)", "E", numeric_token)
                normalized_values.append(numeric_token)
            elif kind == "datetime":
                normalized_values.append(token.strip())
            else:
                text_token = token
                stripped = token.strip()
                if _is_numeric_token(stripped) and re.search(r"[dD][+-]?\d+$", stripped):
                    text_token = re.sub(r"[dD](?=[+-]?\d+$)", "E", stripped)
                normalized_values.append(text_token)
        columns.append(
            DataColumn(
                name=column_name,
                kind=kind,
                values=tuple(normalized_values),
                unit=unit,
                original_name=original_name,
            )
        )
    return columns, warnings


def _is_nonfinite_token(token: str) -> bool:
    lowered = token.casefold()
    if lowered in {"nan", "+nan", "-nan", "inf", "+inf", "-inf", "infinity", "+infinity", "-infinity"}:
        return True
    if _NUMERIC_RE.fullmatch(token):
        try:
            return not math.isfinite(float(token.replace("D", "E").replace("d", "e")))
        except ValueError:
            return False
    return False


def _split_unit(name: str) -> tuple[str, str]:
    for pattern in (_UNIT_PARENS_RE, _UNIT_BRACKETS_RE):
        match = pattern.fullmatch(name)
        if match and match.group(1).strip():
            return match.group(1).strip(), match.group(2).strip()
    return name.strip(), ""


def _unique_name(name: str, used: set[str]) -> str:
    candidate = name
    suffix = 2
    while candidate.casefold() in used:
        candidate = f"{name}_{suffix}"
        suffix += 1
    used.add(candidate.casefold())
    return candidate
