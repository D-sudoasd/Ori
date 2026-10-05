"""Route checks from Git changes without caching validation results."""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

CORE = {f"origin_bridge/{name}.py" for name in ("readers", "planning", "models", "dates", "source_cache")}
BUILD = {
    "spectra_to_origin.py", "data_to_origin_cli.py", "origin_bridge/__main__.py",
    "origin_bridge/worker.py", "origin_bridge/session.py", "spectratoorigin.spec",
    "build_exe.bat", "run.bat",
}
LEAF = {
    "origin_bridge/cli.py": {"test_generic_cli", "test_generic_origin", "test_batch_review"},
    "origin_bridge/mcp_server.py": {"test_mcp_server"},
    "origin_bridge/gui.py": {"test_generic_gui", "test_batch_gui"},
    "origin_bridge/gui_drive.py": {"test_batch_gui", "test_generic_gui"},
    "origin_bridge/batch.py": {"test_batch", "test_batch_gui", "test_batch_origin", "test_batch_review"},
    "origin_bridge/batch_config.py": {"test_batch_gui"},
    "origin_bridge/source_summary.py": {"test_generic_gui", "test_batch_gui"},
    "origin_bridge/exporter.py": {"test_generic_exporter", "test_generic_cli", "test_generic_gui", "test_generic_origin", "test_batch", "test_batch_origin", "test_mcp_server"},
}
BUILD_CHECKS = {"scripts/verify_packaged_worker.py"}
IMAGES = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".bmp", ".tif", ".tiff", ".ico"}
DOCS = {".md", ".rst", ".adoc"}
DOCUMENT_DATA = {"docs/agent-efficiency-results.json", "assets/readme/cover-source.json"}


@dataclass(frozen=True)
class Change:
    path: str
    status: str = "M"


def _normal_path(path: str) -> str:
    return path.replace("\\", "/").removeprefix("./").casefold()


def route_changes(changes: Iterable[Change], *, full: bool = False, fallback: str = "") -> dict[str, object]:
    files = sorted({(_normal_path(c.path), c.status.upper()) for c in changes})
    selected: set[str] = set()
    reasons: set[str] = {v for v in ("explicit-full" if full else "", fallback) if v}
    run_full = full or bool(fallback)
    build = agent = run_full

    for path, status in files:
        suffix, name = Path(path).suffix.casefold(), Path(path).name
        if suffix in IMAGES:
            reasons.add("image-only")
            continue
        if path in DOCUMENT_DATA or suffix in DOCS or name in {"readme", "license"} or name.startswith(("readme.", "license.", "changelog.")):
            reasons.add("documentation-only")
            continue
        if name.startswith("test_") and suffix == ".py":
            if status.startswith("D") or not re.fullmatch(r"test_[A-Za-z0-9_]+", Path(path).stem):
                run_full = True
                reasons.add("test-deleted-or-unimportable")
            else:
                selected.add(Path(path).stem)
            agent |= name == "test_mcp_server.py" and not status.startswith("D")
            continue
        if path == "scripts/check_changes.py":
            selected.add("test_check_changes")
        elif path in BUILD_CHECKS:
            build = True
            reasons.add("packaged-runtime-verifier")
        elif status.startswith(("A", "?")) and suffix == ".py":
            run_full = build = agent = True
            reasons.add("new-production-module")
        elif path in CORE:
            run_full = build = agent = True
            reasons.add("shared-data-contract")
        elif path in BUILD or path.startswith("requirements") or path.startswith(".github/workflows/"):
            run_full = build = True
            agent |= path.startswith("requirements") or path.startswith(".github/workflows/") or path in {
                "origin_bridge/worker.py", "origin_bridge/session.py"
            }
            reasons.add("packaging-or-ci-boundary")
        elif path in LEAF:
            selected.update(LEAF[path])
            agent |= "test_mcp_server" in LEAF[path]
        elif path.startswith("docs/") and suffix == ".txt":
            reasons.add("documentation-only")
            continue
        else:
            run_full = build = agent = True
            reasons.add("unknown-executable-or-config-boundary")

    agent_args = [] if run_full else sorted(selected & {"test_mcp_server"} if agent else set())
    args = ["discover", "-v"] if run_full else sorted(selected - set(agent_args))
    if not args and agent_args:
        reasons.add("agent-tools-only")
    elif not args and not agent and not build and not reasons:
        reasons.add("no-code-check-needed")
    return {
        "mode": "full" if run_full else "targeted" if args or agent_args else "build-only" if build else "none",
        "changed_files": [{"path": p, "status": s} for p, s in files],
        "unittest_args": args,
        "agent_unittest_args": agent_args,
        "agent": agent,
        "build": build,
        "reasons": sorted(reasons),
    }


def _git(*args: str, cwd: str | Path | None = None) -> bytes:
    result = subprocess.run(["git", *args], cwd=cwd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if result.returncode:
        raise RuntimeError(result.stderr.decode("utf-8", "replace").strip())
    return result.stdout


def _parse(data: bytes) -> list[Change]:
    fields, changes, i = data.split(b"\0"), [], 0
    while i + 1 < len(fields):
        changes.append(Change(fields[i + 1].decode("utf-8", "surrogateescape"), fields[i].decode("ascii", "replace")))
        i += 2
    return changes


def changed_files(base: str | None = None, *, cwd: str | Path | None = None) -> list[Change]:
    changes = []
    command = ("diff", "--name-status", "-z", "--no-renames")
    args = [(*command, base, "--") if base is not None else command]
    if base is None:
        args.append(("diff", "--cached", "--name-status", "-z", "--no-renames"))
    for command in args:
        changes.extend(_parse(_git(*command, cwd=cwd)))
    changes.extend(Change(p.decode("utf-8", "surrogateescape"), "??") for p in _git("ls-files", "--others", "--exclude-standard", "-z", cwd=cwd).split(b"\0") if p)
    return sorted({c.path: c for c in changes}.values(), key=lambda c: c.path.casefold())


def _valid_base(base: str, cwd: str | Path | None = None) -> bool:
    return bool(base) and not set(base) <= {"0"} and subprocess.run(
        ["git", "rev-parse", "--verify", "--quiet", f"{base}^{{commit}}"], cwd=cwd,
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    ).returncode == 0


def _github_output(path: str, plan: dict[str, object]) -> None:
    with open(path, "a", encoding="utf-8", newline="\n") as stream:
        stream.write(f"unittest_args={' '.join(plan['unittest_args'])}\n")
        stream.write(f"agent={str(plan['agent']).lower()}\nbuild={str(plan['build']).lower()}\n")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--base", help="diff the worktree against this commit; include untracked files")
    parser.add_argument("--full", action="store_true", help="select the complete test and build path")
    parser.add_argument("--run", action="store_true", help="run selected unittest checks with this Python")
    parser.add_argument("--github-output", help="append route outputs for a GitHub Actions job")
    args = parser.parse_args(argv)
    fallback = ""
    try:
        if args.full:
            changes = []
        elif args.base is not None and not _valid_base(args.base):
            changes, fallback = [], "missing-or-new-push-base"
        else:
            changes = changed_files(args.base)
    except (OSError, RuntimeError):
        changes, fallback = [], "git-state-unavailable"
    plan = route_changes(changes, full=args.full, fallback=fallback)
    plan["base"] = args.base
    print(json.dumps(plan, ensure_ascii=False, indent=2))
    if args.github_output:
        _github_output(args.github_output, plan)
    commands = [plan["unittest_args"]] if plan["unittest_args"] else []
    if plan["agent_unittest_args"]:
        commands.append(plan["agent_unittest_args"])
    for selected in commands if args.run else ():
        result = subprocess.run([sys.executable, "-m", "unittest", *selected], check=False)
        if result.returncode:
            return result.returncode
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
