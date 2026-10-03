"""Opt-in Origin batch readback and same-machine benchmark.

Ordinary unit tests do not start Origin. Run this file only when no Origin
process is open:

    set ORI_BATCH_ORIGIN=1
    py -3 -m unittest test_batch_origin.py

The test refuses an existing Origin process and does not taskkill Origin by
image name. Inputs and outputs stay in the system temp directory.
"""

from __future__ import annotations

import ctypes
import json
import os
import shutil
import sys
import tempfile
import textwrap
import threading
import time
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent
REPORT_PATH = Path(tempfile.gettempdir()) / "ori_batch_origin_report.json"
QA_DIR = Path(tempfile.gettempdir()) / "ori-batch-qa"


def _enabled() -> bool:
    return sys.platform == "win32" and os.environ.get("ORI_BATCH_ORIGIN") == "1"


def _working_set(pid: int) -> int:
    if isinstance(pid, bool) or not isinstance(pid, int) or pid <= 0:
        return 0

    class Counters(ctypes.Structure):
        _fields_ = [
            ("cb", ctypes.c_ulong),
            ("PageFaultCount", ctypes.c_ulong),
            ("PeakWorkingSetSize", ctypes.c_size_t),
            ("WorkingSetSize", ctypes.c_size_t),
            ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
            ("QuotaPagedPoolUsage", ctypes.c_size_t),
            ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
            ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
            ("PagefileUsage", ctypes.c_size_t),
            ("PeakPagefileUsage", ctypes.c_size_t),
        ]

    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    psapi = ctypes.WinDLL("psapi", use_last_error=True)
    kernel.OpenProcess.restype = ctypes.c_void_p
    kernel.OpenProcess.argtypes = [ctypes.c_uint, ctypes.c_int, ctypes.c_uint]
    kernel.CloseHandle.argtypes = [ctypes.c_void_p]
    psapi.GetProcessMemoryInfo.argtypes = [ctypes.c_void_p, ctypes.POINTER(Counters), ctypes.c_ulong]
    psapi.GetProcessMemoryInfo.restype = ctypes.c_int
    handle = kernel.OpenProcess(0x1000, 0, int(pid))
    if not handle:
        handle = kernel.OpenProcess(0x0410, 0, int(pid))
    if not handle:
        return 0
    try:
        info = Counters()
        info.cb = ctypes.sizeof(Counters)
        if not psapi.GetProcessMemoryInfo(handle, ctypes.byref(info), info.cb):
            return 0
        return int(info.WorkingSetSize)
    finally:
        kernel.CloseHandle(handle)


class _PeakSampler:
    def __init__(self, extra_pids):
        self.extra_pids = extra_pids
        self.peak = 0
        self.error = None
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._run, name="ori-batch-memory", daemon=True)

    def __enter__(self):
        self._thread.start()
        return self

    def __exit__(self, exc_type, exc, traceback):
        self._stop.set()
        self._thread.join(timeout=2)
        return False

    def _run(self) -> None:
        while not self._stop.is_set():
            try:
                pids = {os.getpid(), *self.extra_pids()}
                total = sum(_working_set(pid) for pid in pids)
                self.peak = max(self.peak, total)
            except Exception as exc:
                self.error = repr(exc)
                return
            self._stop.wait(0.25)


@unittest.skipUnless(_enabled(), "set ORI_BATCH_ORIGIN=1 on Windows to run the real Origin batch checks")
class RealOriginBatchTests(unittest.TestCase):
    def setUp(self):
        import spectra_to_origin as sto

        self.sto = sto
        if sto.origin_process_running():
            self.skipTest("Origin is already running; the user's session was not closed")
        self.temp = tempfile.TemporaryDirectory(prefix="ori_batch_origin_", dir=tempfile.gettempdir())
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name).resolve()
        self.baseline = set()
        self.owned: set[int] = set()

    def tearDown(self):
        from origin_bridge.session import origin_pids, terminate_pid

        leftover = (origin_pids() - self.baseline) & self.owned
        for pid in sorted(leftover):
            terminate_pid(pid)

    def test_markers_sheets_graphs_and_one_pid(self):
        from origin_bridge.batch import build_batch_request, execute_batch
        from origin_bridge.session import origin_pids, validate_pdf_file

        self.baseline = set(origin_pids())
        specs = []
        folder = self.root / "readback"
        folder.mkdir()
        for marker in (111, 222, 333):
            path = folder / f"m{marker}.csv"
            path.write_text(f"x,y\n0,{marker}\n1,{marker + 1}\n", encoding="utf-8")
            specs.append({
                "source": str(path),
                "source_name": path.name,
                "sheet": path.stem,
                "marker": marker,
                "title": f"graph-{marker}",
            })
        request = build_batch_request(
            [folder],
            self.root / "readback-out",
            format="opju",
            pdf=False,
            plot={"kind": "line", "x": "x", "y": ["y"], "x_label": "x", "y_label": "y"},
            task_overrides=[{"source": item["source"], "plot": {"title": item["title"]}} for item in specs],
        )
        seen: list[list[int]] = []

        def progress(event):
            if event.get("type") == "task" and event.get("status") == "succeeded":
                seen.append(sorted(origin_pids() - self.baseline))

        started = time.perf_counter()
        with _PeakSampler(lambda: origin_pids() - self.baseline) as sampler:
            result = execute_batch(request, progress=progress)
        elapsed = time.perf_counter() - started
        self.assertIsNone(sampler.error, sampler.error)
        self.assertGreater(sampler.peak, 0)
        self.owned.update(result["origin"]["owned_pids"])
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["origin"]["starts"], 1, result["origin"])
        self.assertEqual(result["origin"]["stops"], 1, result["origin"])
        self.assertTrue(seen, "no Origin PID was visible between successful tasks")
        self.assertTrue(all(len(item) == 1 and item == seen[0] for item in seen), seen)
        by_source = {Path(item["source"]).name: item for item in result["tasks"]}
        for spec in specs:
            task = by_source[spec["source_name"]]
            spec["opju"] = task["output"]["path"]
            self.assertEqual(task["resolved_plot"][0]["title"], spec["title"])
            self.assertEqual(task["pdf_status"], "not_applicable")
        self._wait_until_origin_quiescent()
        readback = self._readback(specs)
        report = {
            "kind": "readback",
            "python": sys.version,
            "executable": sys.executable,
            "origin_install_present": Path(r"D:\Program Files\OriginLab\Origin2025").is_dir(),
            "files": len(specs),
            "elapsed_s": elapsed,
            "origin": result["origin"],
            "pid_samples": seen,
            "peak_working_set_bytes": sampler.peak,
            "readback": readback,
        }
        _write_report("readback", report)
        print("BATCH_ORIGIN_READBACK " + json.dumps(report, ensure_ascii=False))
        self.assertEqual(readback["status"], "ok", readback)

        pdf_folder = self.root / "pdf"
        pdf_folder.mkdir()
        roles = [
            ("first", 111, "graph-first"),
            ("middle", 222, "graph-middle"),
            ("last", 333, "graph-last"),
        ]
        role_paths = {}
        for role, marker, title in roles:
            path = pdf_folder / f"{role}.csv"
            path.write_text(f"x,y\n0,{marker}\n1,{marker + 1}\n", encoding="utf-8")
            role_paths[role] = path
        plain = pdf_folder / "plain.csv"
        plain.write_text("x,y\n0,7\n1,8\n", encoding="utf-8")
        multi = pdf_folder / "multi.xlsx"
        _write_multi_workbook(multi)
        pdf_request = build_batch_request(
            [pdf_folder],
            self.root / "pdf-out",
            format="opju",
            pdf=True,
            plot={"kind": "line", "x": "x", "y": ["y"]},
            task_overrides=[
                {"source": role_paths[role], "plot": {"title": title}}
                for role, _marker, title in roles
            ] + [{"source": plain, "plot": {"kind": "none"}}],
        )
        pdf_started = time.perf_counter()
        with _PeakSampler(lambda: origin_pids() - self.baseline) as pdf_sampler:
            pdf_result = execute_batch(pdf_request)
        self.assertIsNone(pdf_sampler.error, pdf_sampler.error)
        self.assertGreater(pdf_sampler.peak, 0)
        self.owned.update(pdf_result["origin"]["owned_pids"])
        self.assertTrue(pdf_result["ok"], pdf_result)
        self.assertEqual(pdf_result["origin"]["starts"], 1, pdf_result["origin"])
        by_pdf = {Path(item["source"]).name: item for item in pdf_result["tasks"]}
        self.assertEqual(by_pdf["plain.csv"]["pdf_status"], "not_applicable")
        self.assertEqual(by_pdf["plain.csv"]["pdfs"], [])
        qa_specs = []
        for role, marker, title in roles:
            task = by_pdf[f"{role}.csv"]
            self.assertEqual(task["pdf_status"], "ok", task)
            self.assertEqual(len(task["pdfs"]), 1)
            info = validate_pdf_file(Path(task["pdfs"][0]["path"]))
            self.assertGreaterEqual(info["pages"], 1)
            qa_specs.append({
                "role": role,
                "opju": task["output"]["path"],
                "source_name": f"{role}.csv",
                "sheet": role,
                "marker": marker,
                "title": title,
                "pdfs": [item["path"] for item in task["pdfs"]],
            })
        multi_task = by_pdf["multi.xlsx"]
        self.assertEqual(multi_task["pdf_status"], "ok", multi_task)
        self.assertEqual(len(multi_task["pdfs"]), 2)
        for item in multi_task["pdfs"]:
            self.assertGreaterEqual(validate_pdf_file(Path(item["path"]))["pages"], 1)
        qa_specs.append({
            "role": "multi",
            "opju": multi_task["output"]["path"],
            "source_name": "multi.xlsx",
            "checks": [
                {"sheet": "multi_Left", "marker": 444, "title": "multi_Left"},
                {"sheet": "multi_Right", "marker": 555, "title": "multi_Right"},
            ],
            "pdfs": [item["path"] for item in multi_task["pdfs"]],
        })
        self._wait_until_origin_quiescent()
        pdf_readback = self._readback(qa_specs)
        kept = _publish_qa(qa_specs)
        pdf_report = {
            "kind": "pdf",
            "elapsed_s": time.perf_counter() - pdf_started,
            "origin": pdf_result["origin"],
            "peak_working_set_bytes": pdf_sampler.peak,
            "qa_dir": str(QA_DIR.resolve()),
            "samples": kept,
            "readback": pdf_readback,
            "tasks": [
                {"source": item["source"], "status": item["status"], "pdf_status": item["pdf_status"], "pdfs": item["pdfs"]}
                for item in pdf_result["tasks"]
            ],
        }
        _write_report("pdf", pdf_report)
        print("BATCH_ORIGIN_PDF " + json.dumps(pdf_report, ensure_ascii=False))
        self.assertEqual(pdf_readback["status"], "ok", pdf_readback)

    def test_benchmark_pdf_off_old_loop_versus_batch(self):
        from origin_bridge.batch import build_batch_request, execute_batch
        from origin_bridge.exporter import execute_import
        from origin_bridge.planning import create_plan, prepare_plan
        from origin_bridge.session import OriginSession, origin_pids

        self.baseline = set(origin_pids())
        folder = self.root / "bench"
        folder.mkdir()
        files = []
        for index, marker in enumerate((1000, 2000, 3000, 4000), start=1):
            path = folder / f"s{index}.csv"
            path.write_text(f"x,y\n0,{marker}\n1,{marker + 1}\n2,{marker + 2}\n", encoding="utf-8")
            files.append(path)
        options = {
            "header": "auto",
            "skip_rows": 0,
            "delimiter": "auto",
            "encoding": "auto",
            "missing_values": [""],
            "formula_policy": "cached",
        }
        plot = {"kind": "line", "x": "x", "y": ["y"], "y_error": "auto", "title": "auto", "x_label": "auto", "y_label": "auto"}

        def extra():
            return origin_pids() - self.baseline

        old_dir = self.root / "old"
        old_dir.mkdir()
        launches_before = OriginSession.launch_count
        closes_before = OriginSession.close_count
        old_started = time.perf_counter()
        with _PeakSampler(extra) as old_sampler:
            for path in files:
                plan = create_plan([path], old_dir / f"{path.stem}.opju", format="opju", options=options)
                for entry in plan["tables"]:
                    entry["plot"]["kind"] = "line"
                    entry["plot"]["x"] = "x"
                    entry["plot"]["y"] = ["y"]
                receipt = execute_import(prepare_plan(plan))
                self.assertTrue(receipt["ok"], receipt)
        old_elapsed = time.perf_counter() - old_started
        self.assertIsNone(old_sampler.error, old_sampler.error)
        self.assertGreater(old_sampler.peak, 0)
        old_starts = OriginSession.launch_count - launches_before
        old_stops = OriginSession.close_count - closes_before
        self._wait_until_origin_quiescent()

        batch_started = time.perf_counter()
        with _PeakSampler(extra) as batch_sampler:
            result = execute_batch(build_batch_request(
                files, self.root / "new", format="opju", pdf=False, plot=plot, read_options=options,
            ))
        batch_elapsed = time.perf_counter() - batch_started
        self.assertIsNone(batch_sampler.error, batch_sampler.error)
        self.assertGreater(batch_sampler.peak, 0)
        self.owned.update(result["origin"]["owned_pids"])
        report = {
            "kind": "benchmark",
            "pdf": False,
            "files": len(files),
            "rows_per_file": 3,
            "columns_per_file": 2,
            "python": sys.version,
            "executable": sys.executable,
            "origin_install_present": Path(r"D:\Program Files\OriginLab\Origin2025").is_dir(),
            "old": {
                "elapsed_s": old_elapsed,
                "origin_starts": old_starts,
                "origin_stops": old_stops,
                "peak_working_set_bytes": old_sampler.peak,
            },
            "batch": {
                "elapsed_s": batch_elapsed,
                "origin_starts": result["origin"]["starts"],
                "origin_stops": result["origin"]["stops"],
                "peak_working_set_bytes": batch_sampler.peak,
                "ok": result["ok"],
            },
        }
        _write_report("benchmark", report)
        print("BATCH_ORIGIN_BENCHMARK " + json.dumps(report, ensure_ascii=False))
        self.assertTrue(result["ok"], result)
        self.assertEqual(old_starts, len(files))
        self.assertEqual(result["origin"]["starts"], 1)
        self._wait_until_origin_quiescent()

    def test_two_graph_snapshot_recovery_keeps_hashes(self):
        import hashlib

        from origin_bridge.batch import build_batch_request, execute_batch
        from origin_bridge.session import origin_pids, validate_pdf_file

        self.baseline = set(origin_pids())
        folder = self.root / "snapshot"
        folder.mkdir()
        source = folder / "two.xlsx"
        _write_multi_workbook(source)
        output = self.root / "snapshot-out"
        plot = {"kind": "line", "x": "x", "y": ["y"]}
        started = time.perf_counter()
        first = execute_batch(build_batch_request([source], output, format="opju", pdf=True, plot=plot))
        self.owned.update(first["origin"]["owned_pids"])
        self.assertTrue(first["ok"], first)
        task = first["tasks"][0]
        self.assertEqual(task["pdf_status"], "ok", task)
        self.assertEqual(len(task["pdfs"]), 2, task)
        opju = Path(task["output"]["path"])
        pdfs = sorted(task["pdfs"], key=lambda item: int(item["index"]))
        kept = Path(pdfs[0]["path"])
        missing = Path(pdfs[1]["path"])

        def digest(path: Path) -> str:
            return hashlib.sha256(path.read_bytes()).hexdigest()

        before = {"opju": digest(opju), "kept": digest(kept), "filled": digest(missing)}
        missing.unlink()
        self._wait_until_origin_quiescent()
        resumed = execute_batch(build_batch_request(
            [source], output, format="opju", pdf=True, plot=plot, resume=True,
        ))
        self.owned.update(resumed["origin"]["owned_pids"])
        resumed_task = resumed["tasks"][0]
        self.assertEqual(resumed_task["pdf_status"], "ok", resumed)
        self.assertEqual(sorted(item["index"] for item in resumed_task["pdfs"]), [1, 2])
        after = {
            "opju": digest(opju),
            "kept": digest(kept),
            "filled": digest(missing),
        }
        self.assertEqual(after["opju"], before["opju"])
        self.assertEqual(after["kept"], before["kept"])
        self.assertGreaterEqual(validate_pdf_file(missing)["pages"], 1)
        self.assertGreaterEqual(validate_pdf_file(kept)["pages"], 1)
        report = {
            "kind": "snapshot-recovery",
            "files": 1,
            "graphs": 2,
            "python": sys.version,
            "executable": sys.executable,
            "elapsed_s": time.perf_counter() - started,
            "origin_first": first["origin"],
            "origin_resume": resumed["origin"],
            "before": before,
            "after": after,
            "opju_unchanged": after["opju"] == before["opju"],
            "kept_pdf_unchanged": after["kept"] == before["kept"],
            "filled_pdf_pages": validate_pdf_file(missing)["pages"],
        }
        report_path = Path(tempfile.gettempdir()) / "ori_batch_r3_snapshot_report.json"
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        print("BATCH_ORIGIN_SNAPSHOT " + json.dumps(report, ensure_ascii=False))
        self._wait_until_origin_quiescent()

    def _wait_until_origin_quiescent(self) -> None:
        from origin_bridge.session import origin_pids, terminate_pid

        deadline = time.monotonic() + 12
        while time.monotonic() < deadline:
            if not (origin_pids() - self.baseline):
                return
            time.sleep(0.25)
        for pid in sorted((origin_pids() - self.baseline) & self.owned):
            terminate_pid(pid)
        deadline = time.monotonic() + 8
        while time.monotonic() < deadline:
            if not ((origin_pids() - self.baseline) & self.owned):
                return
            time.sleep(0.25)

    def _readback(self, specs: list[dict]) -> dict:
        script = textwrap.dedent("""
            import json, sys
            import originpro as op
            import spectra_to_origin as sto
            specs = json.loads(sys.argv[1])
            if sto.origin_process_running():
                raise SystemExit("Origin is already running; readback did not attach")
            op.set_show(False)
            try:
                for spec in specs:
                    opened = op.open(spec["opju"])
                    if opened is False:
                        raise SystemExit("open failed: " + spec["opju"])
                    books = list(op.pages("w"))
                    if len(books) != 1:
                        raise SystemExit("expected one workbook, got %s" % len(books))
                    book = books[0]
                    provenance = []
                    sheets = {}
                    for sheet in book:
                        if sheet.lname == "Import Provenance":
                            for index in range(sheet.cols):
                                provenance.extend(str(value) for value in sheet.to_list(index))
                        else:
                            sheets[sheet.lname] = sheet
                    if spec["source_name"] not in "\\n".join(provenance):
                        raise SystemExit("provenance missing %s" % spec["source_name"])
                    checks = spec.get("checks") or [{
                        "sheet": spec["sheet"], "marker": spec["marker"], "title": spec["title"],
                    }]
                    for check in checks:
                        data_sheet = sheets.get(check["sheet"])
                        if data_sheet is None:
                            raise SystemExit("missing sheet %s in %s" % (check["sheet"], sorted(sheets)))
                        y_values = data_sheet.to_list(1)
                        if float(y_values[0]) != float(check["marker"]):
                            raise SystemExit("marker mismatch %s %s: %s" % (spec["opju"], check["sheet"], y_values))
                    graphs = [graph.lname for graph in op.pages("g")]
                    expected = [check["title"] for check in checks]
                    if graphs != expected:
                        raise SystemExit("graph mismatch %s: %s" % (spec["opju"], graphs))
                print("READBACK_OK")
            finally:
                sto.close_origin_app(op, started=True)
        """)
        import subprocess
        proc = subprocess.run(
            [sys.executable, "-c", script, json.dumps(specs)],
            cwd=ROOT, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=240,
        )
        return {
            "status": "ok" if proc.returncode == 0 and "READBACK_OK" in proc.stdout else "failed",
            "returncode": proc.returncode,
            "stdout": proc.stdout,
            "stderr": proc.stderr,
        }


def _write_multi_workbook(path: Path) -> None:
    from openpyxl import Workbook

    book = Workbook()
    left = book.active
    left.title = "Left"
    left.append(["x", "y"])
    left.append([0, 444])
    left.append([1, 445])
    right = book.create_sheet("Right")
    right.append(["x", "y"])
    right.append([0, 555])
    right.append([1, 556])
    book.save(path)
    book.close()


def _publish_qa(specs: list[dict]) -> list[dict]:
    """Copy the small acceptance set out of the temp test directory.

    The test directory is deleted with the test. This folder stays under the
    user temp directory so a later read-only check can open the same files.
    It is not part of the git tree.
    """
    if QA_DIR.exists():
        shutil.rmtree(QA_DIR)
    QA_DIR.mkdir(parents=True)
    published = []
    for spec in specs:
        role_dir = QA_DIR / spec["role"]
        role_dir.mkdir()
        opju = role_dir / Path(spec["opju"]).name
        shutil.copy2(spec["opju"], opju)
        pdfs = []
        for raw in spec["pdfs"]:
            dest = role_dir / Path(raw).name
            shutil.copy2(raw, dest)
            pdfs.append(str(dest.resolve()))
        published.append({
            "role": spec["role"],
            "opju": str(opju.resolve()),
            "pdfs": pdfs,
            "source_name": spec["source_name"],
        })
    manifest = {
        "qa_dir": str(QA_DIR.resolve()),
        "report": str(REPORT_PATH.resolve()),
        "samples": published,
    }
    (QA_DIR / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return published


def _write_report(section: str, payload: dict) -> None:
    current = {}
    if REPORT_PATH.is_file():
        try:
            current = json.loads(REPORT_PATH.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            current = {}
    if not isinstance(current, dict):
        current = {}
    current[section] = payload
    text = json.dumps(current, ensure_ascii=False, indent=2)
    REPORT_PATH.write_text(text, encoding="utf-8")
    if QA_DIR.is_dir():
        (QA_DIR / "report.json").write_text(text, encoding="utf-8")


if __name__ == "__main__":
    unittest.main()
