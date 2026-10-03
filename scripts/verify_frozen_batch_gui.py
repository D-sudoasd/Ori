"""Run the production GUI executable through a real per-file XLSX batch.

This is not a substitute build. It launches dist/SpectraToOrigin.exe, which
spawns its own batch worker. Set ORI_GUI_BATCH_DRIVE only for that process.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    exe = (root / "dist" / "SpectraToOrigin.exe").resolve()
    cli = (root / "dist" / "DataToOriginCLI.exe").resolve()
    if not exe.is_file() or not cli.is_file():
        print(f"missing packaged executable under {root / 'dist'}", file=sys.stderr)
        return 2
    folder = Path(tempfile.gettempdir()) / "ori-frozen-batch-gui"
    if folder.exists():
        import shutil
        shutil.rmtree(folder)
    folder.mkdir(parents=True)
    try:
        _ = folder
        sources = folder / "sources"
        sources.mkdir()
        paths = []
        for index in range(8):
            path = sources / f"sample{index:02d}.csv"
            path.write_text("x,y\n" + "\n".join(f"{row},{row}" for row in range(400)) + "\n", encoding="utf-8")
            paths.append(path)
        repair = sources / "zzz-bad.csv"
        repair.write_bytes(b"")
        paths.append(repair)
        output = folder / "out"
        output.mkdir()
        shots = folder / "shots"
        shots.mkdir()
        report_path = folder / "report.json"
        spec_path = folder / "drive.json"
        spec_path.write_text(json.dumps({
            "inputs": [str(path) for path in paths],
            "output_dir": str(output),
            "name": "frozen-xlsx",
            "repair_path": str(repair),
            "report": str(report_path),
            "screenshots": {
                "started": str(shots / "started.bmp"),
                "cancelled": str(shots / "cancelled.bmp"),
                "resumed": str(shots / "resumed.bmp"),
                "retried": str(shots / "retried.bmp"),
            },
        }, ensure_ascii=False), encoding="utf-8")
        env = os.environ.copy()
        env["ORI_GUI_BATCH_DRIVE"] = str(spec_path)
        proc = subprocess.run([str(exe)], env=env, cwd=str(folder), timeout=240)
        print(f"gui_exit={proc.returncode}")
        report = json.loads(report_path.read_text(encoding="utf-8"))
        print(json.dumps({key: report[key] for key in report if key != "phases"}, ensure_ascii=False, indent=2))
        if not report.get("ok"):
            return 1
        if report.get("inspect_calls_before_export") != 0:
            print("discovery inspected sources", file=sys.stderr)
            return 1
        if not report.get("child_pids") or report["parent_pid"] in report["child_pids"]:
            print("batch child pid was not distinct", file=sys.stderr)
            return 1
        if int(report.get("window_count_max") or 0) != 1:
            print("unexpected extra window", file=sys.stderr)
            return 1
        for name, path in report.get("screenshots", {}).items():
            size = Path(path).stat().st_size
            print(f"screenshot {name} {path} bytes={size}")
            if size < 1000:
                return 1
        cli_out = folder / "cli-out"
        cli_out.mkdir()
        cli_run = subprocess.run(
            [str(cli), "batch", "-i", str(paths[0]), "-o", str(cli_out), "--format", "xlsx", "--plot", "none"],
            cwd=str(folder),
            capture_output=True,
            timeout=120,
        )
        print(cli_run.stdout.decode("utf-8", errors="replace"))
        if cli_run.returncode not in (0, 4):
            print(cli_run.stderr.decode("utf-8", errors="replace"), file=sys.stderr)
            return 1
        produced = list(cli_out.glob("*.xlsx"))
        print(f"cli_xlsx={produced}")
        if not produced:
            return 1
        print(f"FROZEN_OK exe={exe}")
        print(f"FROZEN_OK cli={cli}")
        print(f"FROZEN_OK report={report_path}")
        return 0
    finally:
        pass


if __name__ == "__main__":
    raise SystemExit(main())
