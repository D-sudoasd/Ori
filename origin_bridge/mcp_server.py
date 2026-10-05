"""Optional MCP stdio server for agent-driven tabular imports into Origin.

Install the optional dependency with ``python -m pip install -r
requirements-agent.txt``, then run ``python -m origin_bridge.mcp_server``.

The server creates reviewable JSON plans and explicitly executes them. Source
inspection and standalone validation are available when needed. Execution can
write an XLSX file or start Origin and write an OPJU project. It never accepts
Python, LabTalk, shell commands, or executable plan fields.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any


class OptionalDependencyError(RuntimeError):
    """Raised when the optional MCP SDK has not been installed."""


def _source_paths(paths: list[str]) -> list[Path]:
    if not isinstance(paths, list) or not paths:
        raise ValueError("paths must be a non-empty list of file or folder paths")
    if any(not isinstance(path, str) or not path.strip() for path in paths):
        raise ValueError("each path must be a non-empty string")
    return [Path(path).expanduser() for path in paths]


def _base_path(base_dir: str | None) -> Path | None:
    if base_dir is None:
        return None
    if not isinstance(base_dir, str) or not base_dir.strip():
        raise ValueError("base_dir must be omitted or a non-empty path")
    return Path(base_dir).expanduser()


def build_server() -> Any:
    """Build and return the MCP SDK server without starting its transport."""
    try:
        from mcp.server import MCPServer
        from mcp.server.mcpserver.exceptions import ToolError
    except ModuleNotFoundError as exc:
        if exc.name == "mcp":
            raise OptionalDependencyError(
                "The optional MCP SDK is not installed. Run "
                "`python -m pip install -r requirements-agent.txt` and retry."
            ) from exc
        raise

    # Import the application layer only after the optional SDK is available.
    from .exporter import execute_import
    from .models import DataImportError
    from .source_cache import SourceCache
    from .planning import (
        create_plan,
        describe_prepared,
        inspect_inputs,
        prepare_plan,
    )

    server = MCPServer("DataToOrigin")
    cache = SourceCache()

    def expected_error(exc: Exception) -> ToolError:
        """Report rejected data and filesystem conditions without a server traceback."""
        return ToolError(str(exc))

    @server.tool()
    def inspect_data(paths: list[str], options: dict[str, Any] | None = None) -> dict[str, Any]:
        """Inspect supported tabular files and report sheets, columns, types and samples.

        This only reads source data. It does not create files or launch Origin.
        Paths are local paths visible to the machine running this MCP server.
        """
        try:
            return inspect_inputs(_source_paths(paths), options, cache=cache)
        except (DataImportError, OSError, ValueError) as exc:
            raise expected_error(exc) from exc

    @server.tool()
    def create_import_plan(
        paths: list[str],
        output_path: str,
        output_format: str = "opju",
        options: dict[str, Any] | None = None,
        overwrite: bool = False,
        keep_open: bool = False,
        include_inspection: bool = False,
    ) -> dict[str, Any]:
        """Create a declarative JSON import plan without writing output or starting Origin.

        output_format is "opju" or "xlsx". Use include_inspection=true to return
        {plan, inspection} with types, samples, ranges and warnings in one call.
        Otherwise returns the plan itself. Review the result, then execute it;
        standalone validation is useful for edits or unresolved mapping questions.
        """
        try:
            if not isinstance(output_path, str) or not output_path.strip():
                raise ValueError("output_path must be a non-empty file path")
            if any(type(value) is not bool for value in (overwrite, keep_open, include_inspection)):
                raise ValueError("overwrite, keep_open and include_inspection must be boolean values")
            sources = _source_paths(paths)
            with cache.operation():
                plan = create_plan(
                    sources, Path(output_path).expanduser(), format=output_format,
                    options=options, overwrite=overwrite, keep_open=keep_open, cache=cache,
                )
                if include_inspection:
                    return {"plan": plan, "inspection": inspect_inputs(sources, options, cache=cache)}
                return plan
        except (DataImportError, OSError, ValueError) as exc:
            raise expected_error(exc) from exc

    @server.tool()
    def validate_import_plan(
        plan: dict[str, Any], base_dir: str | None = None
    ) -> dict[str, Any]:
        """Validate a JSON import plan, source hashes, columns and output settings.

        This performs no output writes and does not launch Origin. base_dir can
        resolve relative source paths from the directory containing a saved plan.
        """
        try:
            prepared = prepare_plan(plan, base_dir=_base_path(base_dir), cache=cache)
            return {"valid": True, "prepared": describe_prepared(prepared)}
        except (DataImportError, OSError, ValueError) as exc:
            raise expected_error(exc) from exc

    @server.tool()
    def execute_import_plan(
        plan: dict[str, Any], confirm: bool, base_dir: str | None = None
    ) -> dict[str, Any]:
        """Execute the supplied, validated JSON plan after explicit confirmation.

        Set confirm=true to authorize the write. XLSX export writes the requested
        workbook. OPJU export may start Origin and writes the requested project.
        Existing output is replaced only when the plan explicitly sets overwrite=true.
        """
        if type(confirm) is not bool or not confirm:
            raise ToolError(
                "Execution was not authorized. Call this tool with confirm=true "
                "after reviewing the plan. Execution validates the final plan."
            )
        try:
            prepared = prepare_plan(plan, base_dir=_base_path(base_dir), cache=cache)
            return execute_import(prepared)
        except (DataImportError, OSError, ValueError) as exc:
            raise expected_error(exc) from exc

    return server


def main() -> int:
    """Start the MCP stdio transport, keeping normal output off its wire."""
    try:
        server = build_server()
    except OptionalDependencyError as exc:
        print(str(exc), file=sys.stderr)
        return 2
    server.run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
