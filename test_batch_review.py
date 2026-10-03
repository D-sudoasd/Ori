"""Review fixes: ownership, record durability, PDF recovery, lock, and explicit modes.

These tests do not start Origin. Process-control checks use injected handles.
"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import tempfile
import threading
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

from openpyxl import Workbook

from origin_bridge.batch import build_batch_request, execute_batch, load_batch_record, log_path_for, read_batch_log
from origin_bridge.cli import build_parser
from origin_bridge.exporter import origin_export_lock
from origin_bridge.models import DataImportError
from origin_bridge.session import (
    OriginSession,
    OwnedProcess,
    _origin_image,
    claim_sentinel_holder,
    pid_creation_filetime,
    sentinel_holders,
    validate_pdf_file,
)
from test_batch import PDF, FakeSession, run_cli

ROOT = Path(__file__).resolve().parent


class ReviewFixTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="ori_batch_review_")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name).resolve()

    def write_csv(self, path: Path, text: str = "x,y\n0,1\n1,2\n") -> Path:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding="utf-8")
        return path

    def workbook(self, path: Path) -> Path:
        path.parent.mkdir(parents=True, exist_ok=True)
        book = Workbook()
        left = book.active
        left.title = "Left"
        left.append(["x", "y"])
        left.append([0, 1])
        right = book.create_sheet("Right")
        right.append(["x", "y"])
        right.append([2, 3])
        book.save(path)
        book.close()
        return path

    def export(self, inputs, output, **kwargs):
        factory = kwargs.pop("session_factory", None)
        progress = kwargs.pop("progress", None)
        cancel = kwargs.pop("cancel_event", None)
        resume = kwargs.pop("resume", None)
        if resume is not None:
            kwargs["resume"] = resume
        request = build_batch_request(list(inputs), output, **kwargs)
        return execute_batch(
            request, progress=progress, cancel_event=cancel, session_factory=factory, resume=resume,
        )

    def snapshots(self, record: Path) -> dict[str, bytes | None]:
        paths = {
            "primary": record,
            "bak": Path(str(record) + ".bak"),
            "previous": record.with_name(record.name + ".previous"),
            "log": log_path_for(record),
        }
        found = {}
        for name, path in paths.items():
            found[name] = path.read_bytes() if path.is_file() else None
        return found

    def test_failed_discovery_keeps_the_previous_success_record(self):
        source = self.write_csv(self.root / "kept.csv", "x,y\n0,7\n")
        output = self.root / "kept-out"
        first = self.export([source], output, format="xlsx", plot={"kind": "none"})
        self.assertTrue(first["ok"], first)
        target = Path(first["tasks"][0]["output"]["path"])
        original = target.read_bytes()
        record = Path(first["record_path"])
        before = self.snapshots(record)
        self.assertIsNotNone(before["primary"])

        missing = self.root / "missing-input"
        empty = self.root / "empty-input"
        empty.mkdir()
        generated_only = self.root / "generated-only"
        generated_only.mkdir()
        planted = generated_only / "keep__0123456789ab.xlsx"
        planted.write_bytes(b"leave-me")
        for inputs, layout in (
            ([missing], "per_file"),
            ([empty], "per_file"),
            ([generated_only], "per_file"),
            ([missing], "unified"),
            ([empty], "unified"),
            ([generated_only], "unified"),
        ):
            with self.assertRaises(DataImportError):
                self.export(inputs, output, layout=layout, format="xlsx", plot={"kind": "none"})
            self.assertEqual(self.snapshots(record), before)
            self.assertEqual(target.read_bytes(), original)
            self.assertEqual(planted.read_bytes(), b"leave-me")

        resumed = self.export([source], output, format="xlsx", plot={"kind": "none"}, resume=True)
        self.assertEqual(resumed["counts"]["skipped"], 1)
        self.assertEqual(target.read_bytes(), original)
        self.assertEqual(Path(resumed["record_path"]).read_bytes(), before["primary"])

    def test_explicit_generated_name_is_imported_and_scans_warn(self):
        explicit = self.workbook(self.root / "named" / "sample__abcdefabcdef.xlsx")
        result = self.export([explicit], self.root / "explicit-out", format="xlsx", plot={"kind": "none"})
        self.assertTrue(result["ok"], result)
        self.assertEqual(len(result["tasks"]), 1)
        self.assertFalse(any("generated batch output name" in warning for warning in result["warnings"]))
        self.assertEqual(explicit.read_bytes()[:2], b"PK")

        folder = self.root / "scan"
        folder.mkdir()
        kept = self.write_csv(folder / "kept.csv", "x,y\n0,4\n")
        generated = folder / "other__abcdefabcdef.xlsx"
        generated.write_bytes(b"generated")
        scanned = self.export([folder], self.root / "scan-out", format="xlsx", plot={"kind": "none"})
        self.assertEqual([Path(item["source"]).name for item in scanned["tasks"]], ["kept.csv"])
        self.assertTrue(any("generated batch output name" in warning for warning in scanned["warnings"]))
        self.assertIn(str(generated.resolve()), "\n".join(scanned["warnings"]))
        self.assertEqual(generated.read_bytes(), b"generated")
        self.assertTrue(kept.is_file())

    def test_pdf_failure_is_visible_and_retry_keeps_the_project(self):
        source = self.write_csv(self.root / "pdf.csv", "x,y\n0,1\n1,2\n")
        output = self.root / "pdf-out"
        plot = {"kind": "line", "x": "x", "y": ["y"], "title": "Only"}
        events = []
        failed = self.export(
            [source], output, format="opju", pdf=True, plot=plot,
            session_factory=lambda: FakeSession(pdf_error=True), progress=events.append,
        )
        task = failed["tasks"][0]
        self.assertEqual(task["status"], "succeeded")
        self.assertIsNone(task["error"])
        self.assertEqual(task["pdf_status"], "failed")
        self.assertEqual(task["pdf_errors"][0]["message"], "pdf boom")
        self.assertTrue(any(event.get("pdf_errors") and event["pdf_errors"][0]["message"] == "pdf boom" for event in events))
        stored = load_batch_record(failed["record_path"])["tasks"][task["task_id"]]
        self.assertIsNone(stored["error"])
        self.assertEqual(stored["pdf_errors"][0]["message"], "pdf boom")
        self.assertTrue(stored["graph_identities"])
        opju = Path(task["output"]["path"])
        original = opju.read_bytes()
        logged = read_batch_log(log_path_for(failed["record_path"]))
        self.assertTrue(any("pdf boom" in json.dumps(event, ensure_ascii=False) for event in logged))

        resumed = self.export(
            [source], output, format="opju", pdf=True, plot=plot, resume=True,
            session_factory=lambda: FakeSession(),
        )
        self.assertEqual(resumed["tasks"][0]["pdf_status"], "ok", resumed)
        self.assertEqual(opju.read_bytes(), original)
        self.assertTrue(resumed["tasks"][0]["pdfs"])

        sessions: list[FakeSession] = []

        def factory():
            session = FakeSession()
            sessions.append(session)
            return session

        opju.write_bytes(original)
        Path(resumed["tasks"][0]["pdfs"][0]["path"]).unlink()
        load_batch_record(resumed["record_path"])
        retried = self.export(
            [source], output, format="opju", pdf=True, plot=plot, resume=True,
            retry_task_ids=[task["task_id"]], session_factory=factory,
        )
        self.assertEqual(retried["tasks"][0]["pdf_status"], "ok", retried)
        self.assertEqual(opju.read_bytes(), original)
        self.assertEqual(len(sessions[0].recovered), 1)
        self.assertIsNone(retried["tasks"][0]["error"])

    def test_partial_pdf_recovery_and_refuses_replaced_inputs(self):
        source = self.workbook(self.root / "two.xlsx")
        output = self.root / "two-out"
        plot = {"kind": "line", "x": "x", "y": ["y"]}
        sessions: list[FakeSession] = []

        def failing():
            session = FakeSession(pdf_fail_indexes={2})
            sessions.append(session)
            return session

        first = self.export(
            [source], output, format="opju", pdf=True, plot=plot, session_factory=failing,
        )
        task = first["tasks"][0]
        self.assertEqual(task["status"], "succeeded")
        self.assertIsNone(task["error"])
        self.assertEqual(task["pdf_status"], "failed")
        self.assertEqual(len(task["pdfs"]), 1)
        self.assertEqual(task["pdfs"][0]["index"], 1)
        self.assertIn("pdf boom", task["pdf_errors"][0]["message"])
        opju = Path(task["output"]["path"])
        original = opju.read_bytes()
        first_pdf = Path(task["pdfs"][0]["path"])
        first_bytes = first_pdf.read_bytes()

        def recovering():
            session = FakeSession()
            sessions.append(session)
            return session

        filled = self.export(
            [source], output, format="opju", pdf=True, plot=plot, resume=True, session_factory=recovering,
        )
        self.assertEqual(filled["tasks"][0]["pdf_status"], "ok", filled)
        self.assertEqual(sorted(item["index"] for item in filled["tasks"][0]["pdfs"]), [1, 2])
        self.assertEqual([item["index"] for item in sessions[-1].recovered[0]], [2])
        self.assertEqual(opju.read_bytes(), original)
        self.assertEqual(first_pdf.read_bytes(), first_bytes)

        before_sessions = len(sessions)
        skipped = self.export(
            [source], output, format="opju", pdf=True, plot=plot, resume=True, session_factory=recovering,
        )
        self.assertEqual(skipped["counts"]["skipped"], 1)
        self.assertEqual(len(sessions), before_sessions)
        self.assertEqual(opju.read_bytes(), original)

        second_pdf = Path(next(item["path"] for item in filled["tasks"][0]["pdfs"] if item["index"] == 2))
        second_bytes = second_pdf.read_bytes()
        second_pdf.unlink()
        deleted = self.export(
            [source], output, format="opju", pdf=True, plot=plot, resume=True, session_factory=recovering,
        )
        self.assertEqual([item["index"] for item in sessions[-1].recovered[0]], [2])
        self.assertEqual(first_pdf.read_bytes(), first_bytes)
        self.assertEqual(opju.read_bytes(), original)
        self.assertEqual(deleted["tasks"][0]["pdf_status"], "ok", deleted)
        self.assertNotEqual(second_pdf.read_bytes(), b"")

        record = Path(filled["record_path"])
        receipt = record.read_bytes()
        stranger = PDF + b"\n% stranger\n"
        second_pdf.write_bytes(stranger)
        before_block = len(sessions)
        blocked = self.export(
            [source], output, format="opju", pdf=True, plot=plot, resume=True, session_factory=recovering,
        )
        self.assertEqual(blocked["counts"]["blocked"], 1)
        self.assertIn("pdf bytes changed", blocked["tasks"][0]["error"]["message"])
        self.assertEqual(record.read_bytes(), receipt)
        self.assertEqual(second_pdf.read_bytes(), stranger)
        self.assertEqual(opju.read_bytes(), original)
        self.assertEqual(load_batch_record(record)["tasks"][task["task_id"]]["status"], "succeeded")
        self.assertEqual(len(sessions), before_block)

        second_pdf.write_bytes(second_bytes)
        changed_book = Workbook()
        left = changed_book.active
        left.title = "Left"
        left.append(["x", "y"])
        left.append([0, 9])
        right = changed_book.create_sheet("Right")
        right.append(["x", "y"])
        right.append([2, 8])
        changed_book.save(source)
        changed_book.close()
        changed = self.export(
            [source], output, format="opju", pdf=True, plot=plot, resume=True, session_factory=recovering,
        )
        self.assertEqual(changed["tasks"][0]["status"], "blocked", changed)
        self.assertEqual(opju.read_bytes(), original)
        self.assertEqual(sessions[-1].recovered, [])

    def test_replaced_project_or_config_is_not_a_pdf_fill(self):
        source = self.write_csv(self.root / "guard.csv", "x,y\n0,1\n1,2\n")
        output = self.root / "guard-out"
        plot = {"kind": "line", "x": "x", "y": ["y"], "title": "Guard"}
        sessions: list[FakeSession] = []

        def factory():
            session = FakeSession()
            sessions.append(session)
            return session

        first = self.export([source], output, format="opju", pdf=True, plot=plot, session_factory=factory)
        opju = Path(first["tasks"][0]["output"]["path"])
        pdf = Path(first["tasks"][0]["pdfs"][0]["path"])
        original = opju.read_bytes()
        pdf_bytes = pdf.read_bytes()
        opju.write_bytes(b"OPJU-REPLACED")
        replaced = self.export(
            [source], output, format="opju", pdf=True, plot=plot, resume=True, session_factory=factory,
        )
        self.assertEqual(replaced["tasks"][0]["status"], "blocked")
        self.assertEqual(opju.read_bytes(), b"OPJU-REPLACED")
        self.assertEqual(pdf.read_bytes(), pdf_bytes)
        self.assertFalse(sessions[-1].recovered)

        opju.write_bytes(original)
        retitled = self.export(
            [source], output, format="opju", pdf=True, plot={**plot, "title": "Other"},
            resume=True, session_factory=factory,
        )
        self.assertEqual(retitled["tasks"][0]["status"], "blocked")
        self.assertEqual(opju.read_bytes(), original)
        self.assertFalse(sessions[-1].recovered)

    def test_recovery_start_failure_keeps_the_success_receipt(self):
        source = self.write_csv(self.root / "dead.csv")
        output = self.root / "dead-out"
        plot = {"kind": "line", "x": "x", "y": ["y"], "title": "Dead"}
        first = self.export(
            [source], output, format="opju", pdf=True, plot=plot,
            session_factory=lambda: FakeSession(pdf_error=True),
        )
        opju = Path(first["tasks"][0]["output"]["path"])
        original = opju.read_bytes()

        class Dead(FakeSession):
            def start(self):
                raise DataImportError("Could not prove which Origin process belongs to this COM instance")

            def export_saved_pdfs(self, project, graphs):
                raise AssertionError("PDF export ran without a proved session")

        failed = self.export(
            [source], output, format="opju", pdf=True, plot=plot, resume=True, session_factory=Dead,
        )
        self.assertEqual(failed["tasks"][0]["status"], "succeeded")
        self.assertIsNone(failed["tasks"][0]["error"])
        self.assertEqual(failed["tasks"][0]["pdf_status"], "failed")
        self.assertIn("Could not prove", failed["tasks"][0]["pdf_errors"][0]["message"])
        self.assertEqual(opju.read_bytes(), original)
        stored = load_batch_record(failed["record_path"])["tasks"][first["tasks"][0]["task_id"]]
        self.assertEqual(stored["status"], "succeeded")
        self.assertEqual(stored["output"]["sha256"], hashlib.sha256(original).hexdigest())
        self.assertIn("Could not prove", stored["pdf_errors"][0]["message"])

    def test_modes_that_cannot_be_honored_are_rejected(self):
        source = self.write_csv(self.root / "mode.csv")
        with self.assertRaisesRegex(DataImportError, "cannot export PDFs"):
            build_batch_request([source], self.root / "u1", layout="unified", format="opju", pdf=True, plot={"kind": "none"})
        with self.assertRaisesRegex(DataImportError, "does not support resume"):
            build_batch_request([source], self.root / "u2", layout="unified", format="xlsx", plot={"kind": "none"}, resume=True)
        with self.assertRaisesRegex(DataImportError, "retry_task_ids"):
            build_batch_request(
                [source], self.root / "u3", format="xlsx", plot={"kind": "none"}, retry_task_ids=["abc"],
            )
        built = build_batch_request([source], self.root / "u4", layout="unified", format="xlsx", plot={"kind": "none"})
        with self.assertRaisesRegex(DataImportError, "does not support resume"):
            execute_batch(built, resume=True)
        per_file = build_batch_request(
            [source], self.root / "u5", format="xlsx", plot={"kind": "none"}, resume=True, retry_task_ids=["abc"],
        )
        per_file["resume"] = False
        with self.assertRaisesRegex(DataImportError, "retry_task_ids"):
            execute_batch(per_file)
        unified_pdf = build_batch_request(
            [source], self.root / "u6", layout="unified", format="opju", plot={"kind": "none"},
        )
        unified_pdf["pdf"] = True
        with self.assertRaisesRegex(DataImportError, "cannot export PDFs"):
            execute_batch(unified_pdf)

        pdf_cli = run_cli(
            "batch", "-i", source, "-o", self.root / "cli-pdf", "--layout", "unified",
            "--format", "opju", "--pdf", "--plot", "none",
        )
        self.assertEqual(pdf_cli.returncode, 4, pdf_cli.stdout + pdf_cli.stderr)
        self.assertIn("cannot export PDFs", pdf_cli.stdout)
        resume_cli = run_cli(
            "batch", "-i", source, "-o", self.root / "cli-resume", "--layout", "unified",
            "--format", "xlsx", "--plot", "none", "--resume",
        )
        self.assertEqual(resume_cli.returncode, 4, resume_cli.stdout + resume_cli.stderr)
        self.assertIn("does not support resume", resume_cli.stdout)

        record = self.root / "retry-record.json"
        from origin_bridge.batch import save_batch_record

        save_batch_record(record, {"tasks": {
            "bad": {"task_id": "bad", "status": "failed", "source": "a"},
            "pdf": {"task_id": "pdf", "status": "succeeded", "pdf_status": "failed", "source": "b"},
        }})
        args = build_parser().parse_args([
            "batch", "-i", str(source), "-o", str(self.root / "cli-retry"),
            "--format", "xlsx", "--plot", "none", "--retry-failed", "--record", str(record),
        ])
        captured = []

        def capture(request, **kwargs):
            del kwargs
            captured.append(request)
            return {"ok": True, "tasks": []}

        with mock.patch("origin_bridge.batch.execute_batch", side_effect=capture):
            from origin_bridge.cli import _run_batch

            payload = _run_batch(args)
        self.assertTrue(payload["ok"])
        self.assertEqual(captured[0]["resume"], True)
        self.assertEqual(captured[0]["retry_task_ids"], ["bad"])

    def test_other_thread_gets_already_active_without_waiting(self):
        started = threading.Event()
        release = threading.Event()
        errors = []

        def hold():
            with origin_export_lock():
                started.set()
                self.assertTrue(release.wait(5))

        def other():
            began = time.perf_counter()
            try:
                with origin_export_lock():
                    errors.append("acquired")
            except Exception as exc:
                errors.append((time.perf_counter() - began, str(exc)))

        holder = threading.Thread(target=hold)
        holder.start()
        self.assertTrue(started.wait(5))
        competitor = threading.Thread(target=other)
        competitor.start()
        competitor.join(2)
        release.set()
        holder.join(5)
        self.assertFalse(competitor.is_alive())
        self.assertEqual(len(errors), 1)
        elapsed, message = errors[0]
        self.assertLess(elapsed, 1.0)
        self.assertIn("already active", message)
        try:
            with origin_export_lock():
                raise RuntimeError("boom")
        except RuntimeError:
            pass
        with origin_export_lock():
            nested = False
            with origin_export_lock():
                nested = True
        self.assertTrue(nested)

    def test_pdf_checker_rejects_objstm_and_accepts_an_uncompressed_page(self):
        plain = self.root / "plain.pdf"
        plain.write_bytes(PDF + b"\n/Filter /FlateDecode\n")
        self.assertGreaterEqual(validate_pdf_file(plain)["pages"], 1)
        objstm = self.root / "objstm.pdf"
        objstm.write_bytes(
            b"%PDF-1.4\n1 0 obj << /Type /ObjStm /N 1 >> stream\nzzzz\nendstream\nendobj\nstartxref\n0\n%%EOF\n"
        )
        with self.assertRaisesRegex(DataImportError, "ObjStm"):
            validate_pdf_file(objstm)
        compressed = self.root / "compressed.pdf"
        compressed.write_bytes(b"%PDF-1.4\n1 0 obj << /Filter /FlateDecode >> stream\nzzzz\nendstream\nendobj\nstartxref\n0\n%%EOF\n")
        with self.assertRaisesRegex(DataImportError, "uncompressed /Type /Page"):
            validate_pdf_file(compressed)

    def test_keep_open_and_recycle_report_every_proved_process(self):
        folder = self.root / "life"
        paths = [self.write_csv(folder / f"{name}.csv") for name in ("a", "b")]
        opened: list[FakeSession] = []

        def factory():
            session = FakeSession()
            session.owned_pids = [81000 + len(opened)]
            session.owned_identities = [{
                "pid": 81000 + len(opened),
                "creation_filetime": 5000 + len(opened),
                "image": "Origin64.exe",
            }]
            opened.append(session)
            return session

        held = self.export(
            [paths[0]], self.root / "held-out", format="opju", plot={"kind": "none"},
            keep_open=True, session_factory=factory,
        )
        self.assertTrue(held["ok"], held)
        self.assertEqual(held["origin"]["stops"], 0)
        self.assertEqual(held["origin"]["owned_pids"], [81000])
        self.assertEqual(held["origin"]["owned_processes"][0]["creation_filetime"], 5000)
        self.assertTrue(opened[0].healthy())
        opened[0].close()

        recycled: list[FakeSession] = []

        def recycled_factory():
            session = FakeSession()
            session.owned_pids = [82000 + len(recycled)]
            session.owned_identities = [{
                "pid": 82000 + len(recycled),
                "creation_filetime": 6000 + len(recycled),
                "image": "Origin64.exe",
            }]
            recycled.append(session)
            return session

        result = self.export(
            [folder], self.root / "recycle-out", format="opju", plot={"kind": "none"},
            recycle_every=1, keep_open=True, session_factory=recycled_factory,
        )
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["origin"]["starts"], 2)
        self.assertEqual(result["origin"]["stops"], 2)
        self.assertGreaterEqual(result["origin"]["recycled"], 1)
        self.assertEqual(result["origin"]["owned_pids"], [82000, 82001])
        self.assertEqual(
            [item["creation_filetime"] for item in result["origin"]["owned_processes"]],
            [6000, 6001],
        )

    def test_unknown_and_reused_pids_are_not_terminated(self):
        if os.name != "nt":
            self.skipTest("Windows process handles")
        session = OriginSession(timeout_s=0)
        session._terminate_handle = lambda handle: (_ for _ in ()).throw(AssertionError("terminated"))
        with mock.patch("origin_bridge.session.origin_pids", return_value={424242}), \
             mock.patch("origin_bridge.session.terminate_pid") as killer:
            self.assertEqual(session.terminate_owned(), [])
            session._deadline = time.monotonic() - 1
            session._enforce_deadline()
            killer.assert_not_called()

        session.started = True
        session._alive = True
        session._live = True
        session.op = object()
        session._owned = [OwnedProcess(42, 1000, 7, r"C:\Origin64.exe")]
        killed = []
        session._terminate_handle = lambda handle: killed.append(handle) or True
        session._handle_status = lambda handle: {"alive": True, "creation": 1000}
        session._pid_creation = lambda pid: 2000
        self.assertFalse(session.healthy())
        self.assertEqual(session.terminate_owned(), [])
        session._enforce_deadline()
        self.assertEqual(killed, [])

        session._pid_creation = lambda pid: 1000
        self.assertTrue(session.healthy())
        with mock.patch("origin_bridge.session.terminate_pid") as killer, \
             mock.patch("origin_bridge.session.origin_pids", return_value={42, 99}):
            session._deadline = time.monotonic() - 1
            session._enforce_deadline()
            killer.assert_not_called()
        self.assertEqual(killed, [7])

        closed = []
        session._close_handle = lambda handle: closed.append(handle)
        session._pid_creation = lambda pid: None
        session.op = None
        session._live = False
        session.close()
        session.close()
        self.assertEqual(closed, [7])

        self.assertIsNone(claim_sentinel_holder([], is_origin_image=lambda pid, image: True))
        self.assertIsNone(claim_sentinel_holder(
            [
                {"pid": 10, "creation_filetime": 1, "image": "Origin64.exe"},
                {"pid": 11, "creation_filetime": 2, "image": "Origin64.exe"},
            ],
            is_origin_image=lambda pid, image: True,
        ))
        claimed = claim_sentinel_holder(
            [{"pid": os.getpid(), "creation_filetime": 5, "image": "friendly name"}],
            is_origin_image=lambda pid, image: image == "friendly name",
        )
        self.assertEqual(claimed["pid"], os.getpid())
        self.assertIsNone(claim_sentinel_holder(
            [{"pid": os.getpid(), "creation_filetime": pid_creation_filetime(os.getpid()) or 1, "image": "Origin"}],
            is_origin_image=_origin_image,
        ))

    def test_restart_manager_sees_this_process_holding_a_file(self):
        if os.name != "nt":
            self.skipTest("Windows Restart Manager")
        directory = Path(tempfile.mkdtemp(prefix="ori_rm_", dir=self.root))
        path = directory / "held.bin"
        path.write_bytes(b"held-by-this-process")
        stream = path.open("rb")
        try:
            holders = sentinel_holders(path)
        finally:
            stream.close()
        ours = [item for item in holders if item["pid"] == os.getpid()]
        self.assertTrue(ours, holders)
        self.assertEqual(ours[0]["creation_filetime"], pid_creation_filetime(os.getpid()))
        self.assertIsNone(claim_sentinel_holder(holders, is_origin_image=_origin_image))

    def test_unproved_launch_does_not_terminate_a_candidate(self):
        class LiveOp:
            def __init__(self):
                self.shown = None
                self.exited = False
                self.value = ""

            def set_show(self, value):
                self.shown = value

            def set_lt_str(self, name, value):
                del name
                self.value = value
                return True

            def get_lt_str(self, name):
                del name
                return self.value

            def lt_exec(self, script):
                del script
                return True

            def lt_int(self, expr):
                del expr
                return 1

            def exit(self):
                self.exited = True

        LiveOp.__module__ = "originpro"
        op = LiveOp()
        closed = []
        support = SimpleNamespace(
            OriginExportError=RuntimeError,
            origin_process_running=lambda: False,
            _load_originpro=lambda: op,
            close_origin_app=lambda app, started: closed.append((app, started)) or app.exit(),
            _cleanup_origin_temp_dirs=lambda directories: None,
        )
        holders = [
            {"pid": 111, "creation_filetime": 10, "image": "Origin64.exe"},
            {"pid": 222, "creation_filetime": 20, "image": "Origin64.exe"},
        ]
        session = OriginSession(timeout_s=0)
        with mock.patch("origin_bridge.exporter._origin_support", return_value=support), \
             mock.patch("origin_bridge.session.origin_pids", return_value=set()), \
             mock.patch("origin_bridge.session.sentinel_holders", return_value=holders), \
             mock.patch("origin_bridge.session.terminate_pid") as killer, \
             mock.patch("origin_bridge.session.terminate_process_handle") as handle_killer:
            with self.assertRaisesRegex(DataImportError, "no candidate process was terminated"):
                session.start()
            killer.assert_not_called()
            handle_killer.assert_not_called()
        self.assertIs(op.shown, False)
        self.assertEqual(closed, [(op, True)])
        self.assertTrue(op.exited)
        self.assertEqual(session.owned_pids, set())


if __name__ == "__main__":
    unittest.main()
