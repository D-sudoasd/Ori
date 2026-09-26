"""JSON command interface shared by people, scripts, and agent tools."""
from __future__ import annotations

import argparse
import contextlib
import json
import sys
from pathlib import Path

from . import __version__
from .models import DataImportError


class _Parser(argparse.ArgumentParser):
    def error(self, message: str) -> None:
        raise DataImportError(message)


def build_parser() -> argparse.ArgumentParser:
    parser = _Parser(description="通用表格 → Origin / Excel；所有操作结果输出 UTF-8 JSON")
    parser.add_argument("--version", action="version", version=__version__)
    subs = parser.add_subparsers(dest="command", required=True)
    for command in ("inspect", "plan", "import"):
        sub = subs.add_parser(command)
        sub.add_argument("-i", "--input", nargs="+", required=True, help="文件或目录，可多个")
        sub.add_argument("--header", choices=("auto", "yes", "no"), default="auto")
        sub.add_argument("--skip-rows", type=int, default=0)
        sub.add_argument("--sheet", help="Excel 工作表名；省略时读取全部非空表")
        sub.add_argument("--delimiter", choices=("auto", "whitespace", "tab", ",", ";"), default="auto")
        sub.add_argument("--encoding", default="auto")
        sub.add_argument("--missing-value", action="append", help="额外的缺失值标记；默认仅空值")
        sub.add_argument("--formula-policy", choices=("cached", "text"), default="cached")
        if command != "inspect":
            sub.add_argument("-o", "--output", required=True)
            sub.add_argument("--format", choices=("opju", "xlsx"), help="默认由输出扩展名确定")
            sub.add_argument("--overwrite", action="store_true", help="允许替换已有输出文件")
            sub.add_argument("--keep-open", action="store_true", help="保存后保持本次 Origin 会话打开")
            sub.add_argument("--plot", choices=("auto", "none", "line", "scatter", "line_symbol", "column"), default="auto")
            sub.add_argument("--x", help="X 列的原始唯一列名；不填时使用建议映射")
            sub.add_argument("--y", nargs="+", help="Y 列的原始唯一列名")
        if command == "plan":
            sub.add_argument("--save", help="将计划另外保存为新 JSON 文件；已有文件不覆盖")
    for command in ("validate", "execute"):
        sub = subs.add_parser(command)
        sub.add_argument("plan", help="导入计划 JSON 文件，相对源路径以计划所在目录为基准")
    return parser


def _read_options(args: argparse.Namespace) -> dict:
    options = {
        "header": {"auto": "auto", "yes": True, "no": False}[args.header],
        "skip_rows": args.skip_rows,
        "delimiter": "\t" if args.delimiter == "tab" else args.delimiter,
        "encoding": args.encoding,
        "formula_policy": args.formula_policy,
    }
    if args.sheet is not None:
        options["sheet"] = args.sheet
    if args.missing_value is not None:
        options["missing_values"] = ["", *args.missing_value]
    return options


def _dispatch(args: argparse.Namespace) -> dict:
    from .planning import create_plan, describe_prepared, inspect_inputs, prepare_plan

    if args.command in ("validate", "execute"):
        plan_path = Path(args.plan).resolve()
        plan = json.loads(plan_path.read_text(encoding="utf-8-sig"))
        prepared = prepare_plan(plan, base_dir=plan_path.parent)
        if args.command == "validate":
            return {"ok": True, **describe_prepared(prepared)}
    else:
        paths = [Path(item) for item in args.input]
        options = _read_options(args)
        if args.command == "inspect":
            return {"ok": True, **inspect_inputs(paths, options=options)}
        output = Path(args.output).resolve()
        output_format = args.format or output.suffix.lstrip(".").lower()
        plan = create_plan(paths, output, format=output_format, options=options,
                           overwrite=args.overwrite, keep_open=args.keep_open)
        for entry in plan["tables"]:
            if args.plot != "auto":
                entry["plot"]["kind"] = args.plot
            if args.plot == "none":
                entry["plot"].update(x=None, y=[], y_error={})
            else:
                if args.x is not None:
                    entry["plot"]["x"] = args.x
                if args.y is not None:
                    entry["plot"]["y"] = args.y
        prepared = prepare_plan(plan)
        if args.command == "plan":
            if args.save:
                target = Path(args.save).resolve()
                if target == prepared.output or any(target == item.table.source for item in prepared.tables):
                    raise DataImportError("计划文件不能与输入或输出文件相同")
                target.parent.mkdir(parents=True, exist_ok=True)
                with target.open("x", encoding="utf-8", newline="\n") as stream:
                    json.dump(plan, stream, ensure_ascii=False, indent=2, allow_nan=False)
                    stream.write("\n")
            # The plan itself is directly reusable; don't wrap it in a result envelope.
            return plan
    from .exporter import execute_import

    return execute_import(prepared)


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    try:
        args = build_parser().parse_args(argv)
        # Third-party libraries must not mix diagnostics into the JSON result.
        with contextlib.redirect_stdout(sys.stderr):
            result = _dispatch(args)
        code = 0
    except Exception as exc:
        result = {"ok": False, "error": {"type": type(exc).__name__, "message": str(exc)}}
        code = 4
    print(json.dumps(result, ensure_ascii=False, indent=2, allow_nan=False))
    return code
