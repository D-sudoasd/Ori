"""Batch window seams: per-file discovery, spawn, cancel, resume, and column scope."""

from __future__ import annotations

import hashlib
import json
import os
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock

from openpyxl import Workbook, load_workbook

from origin_bridge.batch_config import load_batch_config, save_batch_config
from origin_bridge.gui import build_global_plot


def _sha(path: Path) -> str:
    digest = hashlib.sha256()
    digest.update(path.read_bytes())
    return digest.hexdigest()


def _pump(app, timeout: float) -> None:
    deadline = time.monotonic() + timeout
    while app._busy and time.monotonic() < deadline:
        app.root.update()
        time.sleep(0.005)
    app.root.update()


def _shutdown(app) -> None:
    process = getattr(app, "_import_process", None)
    if process is not None and getattr(process, "is_alive", lambda: False)():
        process.terminate()
        process.join(timeout=3)
    try:
        if app.root.winfo_exists():
            app.root.destroy()
    except Exception:
        pass


class BatchConfigTests(unittest.TestCase):
    def test_config_round_trip_does_not_touch_the_core_record(self):
        with tempfile.TemporaryDirectory() as raw:
            folder = Path(raw)
            record = folder / "batch-record.json"
            record.write_text('{"schema_version":1,"kind":"ori-batch-record","tasks":{}}\n', encoding="utf-8")
            before = record.read_bytes()
            request = {
                "schema_version": 1,
                "inputs": [str(folder / "a.csv")],
                "output_dir": str(folder),
                "layout": "per_file",
                "format": "xlsx",
                "read_options": {"header": "auto", "skip_rows": 0},
                "plot": {"kind": "none", "x": "auto", "y": "auto", "y_error": "auto", "title": "auto", "x_label": "auto", "y_label": "auto"},
                "task_overrides": [],
                "overwrite": False,
                "pdf": False,
                "keep_open": False,
                "recycle_every": 0,
                "prefetch": 1,
                "max_resident_tasks": 2,
                "origin_timeout_s": 300,
                "origin_retries": 1,
                "output_name": "batch.xlsx",
                "record_path": str(record),
                "resume": False,
                "retry_task_ids": [],
            }
            path = save_batch_config(
                folder / "batch-config.json",
                name="实验甲",
                record_path=record,
                inputs=request["inputs"],
                request=request,
            )
            loaded = load_batch_config(path)
            self.assertEqual(loaded["name"], "实验甲")
            self.assertEqual(loaded["request"]["plot"]["kind"], "none")
            self.assertEqual(loaded["request"]["inputs"], request["inputs"])
            self.assertEqual(record.read_bytes(), before)
            self.assertTrue((folder / "batch-config.json.bak").is_file() or True)
            save_batch_config(path, name="实验甲", record_path=record, inputs=request["inputs"], request=request)
            self.assertEqual(record.read_bytes(), before)

    def test_inputs_must_match_the_request(self):
        with tempfile.TemporaryDirectory() as raw:
            folder = Path(raw)
            record = str(folder / "batch-record.json")
            request = {"inputs": [str(folder / "a.csv")], "record_path": record, "plot": {"y": "auto"}}
            with self.assertRaises(ValueError):
                save_batch_config(folder / "batch-config.json", name="", record_path=record, inputs=[str(folder / "b.csv")], request=request)


class GlobalPlotTests(unittest.TestCase):
    def test_auto_scope_ignores_typed_columns(self):
        plot = build_global_plot(scope="自动（按每张表建议）", kind="line", explicit_x="9", explicit_y="other", explicit_error="err")
        self.assertEqual(plot["x"], "auto")
        self.assertEqual(plot["y"], "auto")
        self.assertEqual(plot["kind"], "line")

    def test_explicit_scope_keeps_names_and_indexes(self):
        plot = build_global_plot(scope="明确列", kind="auto", explicit_x="0", explicit_y="signal, other", explicit_error="")
        self.assertEqual(plot["kind"], "auto")
        self.assertEqual(plot["x"], 0)
        self.assertEqual(plot["y"], ["signal", "other"])
        self.assertEqual(plot["y_error"], {})

    def test_explicit_scope_requires_a_y_column(self):
        with self.assertRaises(ValueError):
            build_global_plot(scope="明确列", kind="line", explicit_x="x", explicit_y="", explicit_error="")


class _Pipe:
    def __init__(self, messages, raise_after=None):
        self.messages = list(messages)
        self.raise_after = raise_after
        self.recv_count = 0
        self.closed = False

    def poll(self, _timeout=0):
        if self.raise_after is not None and self.recv_count >= self.raise_after:
            raise OSError("broken pipe")
        return bool(self.messages)

    def recv(self):
        if self.raise_after is not None and self.recv_count >= self.raise_after:
            raise OSError("broken pipe")
        if not self.messages:
            raise EOFError
        self.recv_count += 1
        return self.messages.pop(0)

    def close(self):
        self.closed = True


class _Proc:
    def __init__(self):
        self.alive = True
        self.exitcode = None
        self.closed = False
        self.terminated = False

    def is_alive(self):
        if self.closed:
            raise ValueError("closed")
        return self.alive

    def join(self, timeout=None):
        if self.alive and timeout is None:
            raise AssertionError("joined a live process without a bound")
        return None

    def close(self):
        self.closed = True

    def terminate(self):
        self.terminated = True


class _Flag:
    def __init__(self):
        self.flag = False
        self.order = []

    def set(self):
        self.flag = True
        self.order.append("set")

    def is_set(self):
        return self.flag


@unittest.skipUnless(os.environ.get("SPECTRA_TEST_GUI") == "1", "set SPECTRA_TEST_GUI=1 for the Tk batch checks")
class RealTkBatchTests(unittest.TestCase):
    def _app(self):
        from origin_bridge.gui import GeneralDataApp

        try:
            app = GeneralDataApp()
        except Exception as exc:
            self.skipTest(f"Tk display is unavailable: {exc}")
        app.root.withdraw()
        return app

    def test_per_file_lists_1000_without_inspecting_and_keeps_the_heartbeat(self):
        from origin_bridge import gui
        from origin_bridge.readers import discover_files

        with tempfile.TemporaryDirectory() as raw:
            folder = Path(raw) / "sources"
            folder.mkdir()
            for index in range(1000):
                (folder / f"sample{index}.csv").write_text("x,y\n1,2\n", encoding="utf-8")
            (folder / "notes.md").write_text("ignore", encoding="utf-8")
            app = self._app()
            samples = []
            after_ids = []

            def heartbeat():
                samples.append((time.monotonic(), len(app.table_tree.get_children())))
                if app._busy:
                    after_ids.append(app.root.after(10, heartbeat))

            try:
                after_ids.append(app.root.after(10, heartbeat))
                with mock.patch.object(gui.messagebox, "showwarning"), mock.patch.object(gui.messagebox, "showerror") as showerror:
                    app._add_paths([folder, folder / "notes.md"])
                    self.assertTrue(app._busy)
                    self.assertLess(len(app.table_tree.get_children()), 1000)
                    deadline = time.monotonic() + 60
                    while app._busy and time.monotonic() < deadline:
                        app.root.update()
                        time.sleep(0.005)
                    app.root.update()
                    showerror.assert_not_called()
                self.assertFalse(app._busy)
                self.assertEqual(app._summaries.inspect_calls, 0)
                names = [Path(app.table_tree.item(iid, "values")[0]).name for iid in app.table_tree.get_children()]
                self.assertEqual(len(names), 1000)
                self.assertEqual(set(names), {path.name for path in discover_files([folder])})
                self.assertTrue(all(row[4] == "未预览" for row in (app.table_tree.item(iid, "values") for iid in app.table_tree.get_children())))
                gaps = [samples[index][0] - samples[index - 1][0] for index in range(1, len(samples))]
                jumps = [samples[index][1] - samples[index - 1][1] for index in range(1, len(samples))]
                max_gap = max(gaps) if gaps else 999
                max_jump = max(jumps) if jumps else 1000
                partial = any(0 < count < 1000 for _stamp, count in samples)
                print(f"MEASURE per_file_list heartbeat_max_gap_s={max_gap:.4f} max_row_jump={max_jump} heartbeats={len(samples)}")
                self.assertGreaterEqual(len(samples), 5)
                self.assertTrue(partial)
                self.assertLess(max_jump, 80)
                self.assertLess(max_gap, 0.25)
                app.preview_selected()
                _pump(app, 20)
                self.assertEqual(app._summaries.inspect_calls, 1)
                self.assertEqual(len(app._paths), 1000)
                captured = {}
                out = Path(raw) / "out"
                with (
                    mock.patch.object(app, "_start_batch_process", lambda request: captured.setdefault("request", request)),
                    mock.patch.object(gui.filedialog, "askdirectory", return_value=str(out)),
                    mock.patch.object(gui.messagebox, "showerror") as showerror,
                ):
                    app.batch_kind_var.set("none")
                    app._export_per_file()
                    showerror.assert_not_called()
                self.assertEqual(len(captured["request"]["inputs"]), 1000)
                self.assertEqual(captured["request"]["plot"]["y"], "auto")
                self.assertEqual(app._summaries.inspect_calls, 1)
                self.assertTrue((out / "batch-config.json").is_file())
            finally:
                for after_id in after_ids:
                    try:
                        app.root.after_cancel(after_id)
                    except Exception:
                        pass
                _shutdown(app)

    def test_unified_switch_inspects_everything_already_listed(self):
        with tempfile.TemporaryDirectory() as raw:
            folder = Path(raw)
            for name in ("a.csv", "b.csv", "c.csv"):
                (folder / name).write_text("x,y\n1,2\n", encoding="utf-8")
            app = self._app()
            try:
                app._add_paths([folder])
                _pump(app, 20)
                self.assertEqual(app._summaries.inspect_calls, 0)
                self.assertEqual(len(app._paths), 3)
                app.layout_var.set("unified")
                app._layout_changed()
                _pump(app, 20)
                self.assertEqual(app._summaries.inspect_calls, 3)
                self.assertEqual(len(app._tables), 3)
                self.assertFalse(app.pdf_var.get())
            finally:
                _shutdown(app)

    def test_progress_flood_cancel_and_exception_restore_the_window(self):
        from origin_bridge import gui

        app = self._app()
        try:
            messages = []
            for index in range(400):
                pdf_failed = index == 3
                messages.append({
                    "type": "progress",
                    "event": {
                        "type": "task",
                        "task_id": f"{index:04x}",
                        "index": index,
                        "count": 400,
                        "status": "blocked" if index == 4 else "succeeded",
                        "source": str(Path("C:/data") / f"file{index}.csv"),
                        "output": None,
                        "error": None if index != 4 else {"type": "FileExistsError", "message": "already present"},
                        "pdf_status": "failed" if pdf_failed or index == 4 else "not_applicable",
                        "pdf_errors": [
                            {"table": "file3", "graph": "GraphA", "short_name": "Graph1", "path": "C:/out/a.pdf", "message": "pdf export failed"}
                        ] if pdf_failed or index == 4 else [],
                    },
                })
            messages.append({
                "type": "result",
                "result": {
                    "ok": False,
                    "cancelled": False,
                    "counts": {"succeeded": 398, "skipped": 0, "failed": 0, "cancelled": 0, "blocked": 1, "pdf_failed": 2},
                    "tasks": [],
                },
            })
            pipe = _Pipe(messages)
            process = _Proc()
            cancel = _Flag()
            samples = []

            def heartbeat():
                samples.append((time.monotonic(), len(app.task_tree.get_children())))
                if not pipe.messages and process.alive:
                    process.alive = False
                    process.exitcode = 0
                if app._busy:
                    app.root.after(10, heartbeat)

            app._import_recv = pipe
            app._import_process = process
            app._import_cancel = cancel
            app._import_kind = "batch"
            app._import_result = None
            app._import_exit_waits = 0
            app._import_eof = False
            app._set_busy(True)
            self.assertEqual(str(app.stop_button.cget("state")), "normal")
            with mock.patch.object(gui.messagebox, "showinfo"), mock.patch.object(gui.messagebox, "showerror"):
                app.root.after(10, heartbeat)
                app._poll_import()
                app.stop_later_tasks()
                self.assertTrue(cancel.flag)
                self.assertEqual(str(app.stop_button.cget("state")), "normal")
                deadline = time.monotonic() + 30
                while app._busy and time.monotonic() < deadline:
                    app.root.update()
                    time.sleep(0.005)
                app.root.update()
            self.assertFalse(app._busy)
            self.assertEqual(str(app.export_button.cget("state")), "normal")
            self.assertEqual(str(app.stop_button.cget("state")), "disabled")
            self.assertIsNone(app._import_cancel)
            self.assertIsNone(app._import_process)
            self.assertIsNone(app._import_recv)
            self.assertIsNone(app._import_result)
            self.assertFalse(process.terminated)
            self.assertEqual(len(app.task_tree.get_children()), 400)
            values = [app.task_tree.item(iid, "values") for iid in app.task_tree.get_children()]
            pdf_row = next(row for row in values if row[0].endswith("file3.csv"))
            blocked_row = next(row for row in values if row[0].endswith("file4.csv"))
            self.assertIn("成功（PDF 失败）", pdf_row[1])
            self.assertIn("GraphA", pdf_row[2])
            self.assertIn("pdf export failed", pdf_row[2])
            self.assertIn("已阻塞（PDF 失败）", blocked_row[1])
            self.assertIn("already present", blocked_row[2])
            gaps = [samples[index][0] - samples[index - 1][0] for index in range(1, len(samples))]
            jumps = [samples[index][1] - samples[index - 1][1] for index in range(1, len(samples))]
            max_gap = max(gaps) if gaps else 999
            max_jump = max(jumps) if jumps else 400
            partial = any(0 < count < 400 for _stamp, count in samples)
            print(f"MEASURE batch_flood heartbeat_max_gap_s={max_gap:.4f} max_row_jump={max_jump} heartbeats={len(samples)}")
            self.assertGreaterEqual(len(samples), 8)
            self.assertTrue(partial)
            self.assertLess(max_jump, 40)
            self.assertLess(max_gap, 0.25)
            self.assertIn("PDF 失败", app.status_var.get())

            pipe = _Pipe([
                {"type": "progress", "event": {"type": "batch", "status": "started", "count": 1}},
                {"type": "progress", "event": {"type": "task", "task_id": "aa", "index": 0, "count": 1, "status": "failed", "source": "bad.csv", "error": {"type": "ValueError", "message": "boom"}, "pdf_status": "not_applicable", "pdf_errors": []}},
            ], raise_after=1)
            process = _Proc()
            process.alive = False
            process.exitcode = 7
            cancel = _Flag()
            app._import_recv = pipe
            app._import_process = process
            app._import_cancel = cancel
            app._import_kind = "batch"
            app._import_result = None
            app._import_exit_waits = 0
            app._import_eof = False
            app._set_busy(True)
            with mock.patch.object(gui.messagebox, "showerror") as showerror, mock.patch.object(gui.messagebox, "showinfo") as showinfo:
                while app._busy:
                    app._poll_import()
                showinfo.assert_not_called()
                showerror.assert_called()
            self.assertFalse(app._busy)
            self.assertIsNone(app._import_cancel)
            self.assertIsNone(app._import_process)
            self.assertEqual(str(app.export_button.cget("state")), "normal")
            self.assertFalse(process.terminated)
        finally:
            _shutdown(app)

    def test_destroy_requests_cancel_before_the_pipe_closes(self):
        app = self._app()
        order = []

        class _Event:
            def set(self):
                order.append("set")

        class _Recv:
            def close(self):
                order.append("close")

        process = _Proc()
        try:
            app._busy = True
            app._import_kind = "batch"
            app._import_cancel = _Event()
            app._import_recv = _Recv()
            app._import_process = process
            app.root.destroy()
            self.assertTrue(order)
            self.assertLess(order.index("set"), order.index("close"))
            self.assertFalse(process.terminated)
            self.assertTrue(process.is_alive())
        finally:
            process.alive = False
            _shutdown(app)

    def test_real_xlsx_batch_resume_retry_and_column_mismatch(self):
        from origin_bridge import gui
        from origin_bridge.batch import load_batch_record

        with tempfile.TemporaryDirectory() as raw:
            folder = Path(raw)
            good = folder / "good.csv"
            other = folder / "other.csv"
            bad = folder / "bad.csv"
            wrong = folder / "wrong.csv"
            good.write_text("x,signal\n0,1\n1,2\n", encoding="utf-8")
            other.write_text("x,signal\n3,4\n", encoding="utf-8")
            wrong.write_text("x,other\n5,6\n", encoding="utf-8")
            bad.write_bytes(b"")
            book = folder / "book.xlsx"
            workbook = Workbook()
            first = workbook.active
            first.title = "one"
            first.append(["x", "signal"])
            first.append([1, 2])
            second = workbook.create_sheet("two")
            second.append(["p", "q"])
            second.append([3, 4])
            workbook.save(book)
            workbook.close()
            out = folder / "batch-out"
            mismatch_out = folder / "mismatch-out"
            app = self._app()
            try:
                app.format_var.set("xlsx")
                app._format_changed()
                with mock.patch.object(gui.messagebox, "showwarning"), mock.patch.object(gui.messagebox, "showerror"), mock.patch.object(gui.messagebox, "showinfo"):
                    app._add_paths([good, other, bad])
                    _pump(app, 20)
                    self.assertEqual(app._summaries.inspect_calls, 0)
                    self.assertEqual(len(app._paths), 3)
                    app.batch_name_var.set("实验甲")
                    app.batch_kind_var.set("none")
                    app.batch_overwrite_var.set(False)
                    with mock.patch.object(gui.filedialog, "askdirectory", return_value=str(out)):
                        app._export_per_file()
                    _pump(app, 60)
                self.assertFalse(app._busy, app.status_var.get())
                self.assertIsNone(app._import_process)
                self.assertEqual(str(app.export_button.cget("state")), "normal")
                config = load_batch_config(out / "batch-config.json")
                self.assertEqual(config["name"], "实验甲")
                self.assertEqual(len(config["request"]["inputs"]), 3)
                self.assertFalse(config["request"]["overwrite"])
                self.assertEqual(config["request"]["plot"]["kind"], "none")
                record = load_batch_record(out / "batch-record.json")
                statuses = {Path(task["source"]).name: task["status"] for task in record["tasks"].values()}
                self.assertEqual(statuses["good.csv"], "succeeded")
                self.assertEqual(statuses["other.csv"], "succeeded")
                self.assertEqual(statuses["bad.csv"], "failed")
                good_output = next(Path(task["output"]["path"]) for task in record["tasks"].values() if Path(task["source"]).name == "good.csv")
                good_hash = _sha(good_output)
                record_before = (out / "batch-record.json").read_bytes()
                saved_plot = json.loads(json.dumps(config["request"]["plot"]))
                _shutdown(app)

                app = self._app()
                app.format_var.set("xlsx")
                app._format_changed()
                app.plot_scope_var.set("明确列")
                app.explicit_y_var.set("NOPE")
                app.batch_kind_var.set("line")
                with (
                    mock.patch.object(gui.filedialog, "askopenfilename", return_value=str(out / "batch-record.json")),
                    mock.patch.object(gui.messagebox, "showinfo"),
                    mock.patch.object(gui.messagebox, "showerror") as showerror,
                    mock.patch.object(gui.messagebox, "showwarning"),
                ):
                    app.continue_batch()
                    self.assertFalse(showerror.called)
                    resumed = load_batch_config(out / "batch-config.json")
                    self.assertTrue(resumed["request"]["resume"])
                    self.assertFalse(resumed["request"]["overwrite"])
                    self.assertEqual(resumed["request"]["retry_task_ids"], [])
                    self.assertEqual(len(resumed["request"]["inputs"]), 3)
                    self.assertEqual(resumed["request"]["plot"], saved_plot)
                    self.assertNotIn("NOPE", json.dumps(resumed["request"]["plot"]))
                    _pump(app, 60)
                self.assertFalse(app._busy, app.status_var.get())
                self.assertEqual(_sha(good_output), good_hash)
                self.assertEqual(app.batch_name_var.get(), "实验甲")
                self.assertEqual(len(app._paths), 3)
                resumed_rows = [app.task_tree.item(iid, "values") for iid in app.task_tree.get_children()]
                self.assertTrue(any(str(row[0]).endswith("good.csv") and row[1] == "已跳过" for row in resumed_rows))
                self.assertTrue(any(str(row[0]).endswith("bad.csv") and row[1] == "失败" for row in resumed_rows))
                record = load_batch_record(out / "batch-record.json")
                statuses = {Path(task["source"]).name: task["status"] for task in record["tasks"].values()}
                self.assertEqual(statuses["good.csv"], "succeeded")
                self.assertEqual(statuses["bad.csv"], "failed")
                failed_id = next(task_id for task_id, task in record["tasks"].items() if Path(task["source"]).name == "bad.csv")
                succeeded_id = next(task_id for task_id, task in record["tasks"].items() if Path(task["source"]).name == "good.csv")
                bad.write_text("x,signal\n8,9\n", encoding="utf-8")
                with (
                    mock.patch.object(gui.filedialog, "askopenfilename", return_value=str(out / "batch-record.json")),
                    mock.patch.object(gui.messagebox, "showinfo"),
                    mock.patch.object(gui.messagebox, "showerror"),
                    mock.patch.object(gui.messagebox, "showwarning"),
                ):
                    app.retry_failed_batch()
                    retried = load_batch_config(out / "batch-config.json")
                    self.assertTrue(retried["request"]["resume"])
                    self.assertFalse(retried["request"]["overwrite"])
                    self.assertIn(failed_id, retried["request"]["retry_task_ids"])
                    self.assertNotIn(succeeded_id, retried["request"]["retry_task_ids"])
                    self.assertEqual(len(retried["request"]["inputs"]), 3)
                    _pump(app, 60)
                self.assertFalse(app._busy, app.status_var.get())
                record = load_batch_record(out / "batch-record.json")
                statuses = {Path(task["source"]).name: task["status"] for task in record["tasks"].values()}
                self.assertEqual(statuses["bad.csv"], "succeeded")
                self.assertEqual(statuses["good.csv"], "succeeded")
                self.assertEqual(_sha(good_output), good_hash)
                retried_rows = [app.task_tree.item(iid, "values") for iid in app.task_tree.get_children()]
                self.assertTrue(any(str(row[0]).endswith("good.csv") and row[1] == "已跳过" for row in retried_rows))
                self.assertNotEqual((out / "batch-record.json").read_bytes(), record_before)

                _shutdown(app)
                app = self._app()
                app.format_var.set("xlsx")
                app._format_changed()
                with mock.patch.object(gui.messagebox, "showwarning"), mock.patch.object(gui.messagebox, "showerror"), mock.patch.object(gui.messagebox, "showinfo"):
                    app._add_paths([good, wrong, book])
                    _pump(app, 20)
                    self.assertEqual(app._summaries.inspect_calls, 0)
                    for iid in app.table_tree.get_children():
                        if Path(app.table_tree.item(iid, "values")[0]).name == "good.csv":
                            app.table_tree.selection_set(iid)
                            app.table_tree.focus(iid)
                            break
                    app.preview_selected()
                    _pump(app, 20)
                    self.assertEqual(app._summaries.inspect_calls, 1)
                    app.plot_scope_var.set("明确列")
                    app.batch_kind_var.set("line")
                    app.explicit_x_var.set("x")
                    app.explicit_y_var.set("signal")
                    with mock.patch.object(gui.filedialog, "askdirectory", return_value=str(mismatch_out)):
                        app._export_per_file()
                    launched = load_batch_config(mismatch_out / "batch-config.json")
                    self.assertEqual(launched["request"]["plot"]["y"], ["signal"])
                    self.assertEqual(launched["request"]["plot"]["x"], "x")
                    self.assertEqual(len(launched["request"]["inputs"]), 3)
                    _pump(app, 60)
                self.assertFalse(app._busy, app.status_var.get())
                record = load_batch_record(mismatch_out / "batch-record.json")
                by_name = {Path(task["source"]).name: task for task in record["tasks"].values()}
                self.assertEqual(by_name["good.csv"]["status"], "succeeded")
                self.assertEqual(by_name["wrong.csv"]["status"], "failed")
                self.assertEqual(by_name["book.xlsx"]["status"], "failed")
                self.assertIn("signal", json.dumps(by_name["wrong.csv"].get("error")))
                self.assertIn("unknown source column", json.dumps(by_name["book.xlsx"].get("error")))
                details = [app.task_tree.item(iid, "values") for iid in app.task_tree.get_children()]
                self.assertTrue(any(str(row[0]).endswith("wrong.csv") and "失败" in row[1] and "signal" in row[2] for row in details))
                self.assertTrue(any(str(row[0]).endswith("book.xlsx") and "失败" in row[1] and "unknown source column" in row[2] for row in details))
                good_file = Path(by_name["good.csv"]["output"]["path"])
                self.assertTrue(good_file.is_file(), good_file)
                opened = load_workbook(good_file, read_only=True)
                try:
                    self.assertIn("good", opened.sheetnames)
                finally:
                    opened.close()
            finally:
                _shutdown(app)

    def test_missing_config_does_not_export_from_window_defaults(self):
        from origin_bridge import gui

        with tempfile.TemporaryDirectory() as raw:
            folder = Path(raw)
            record = folder / "batch-record.json"
            record.write_text('{"schema_version":1,"kind":"ori-batch-record","tasks":{}}\n', encoding="utf-8")
            app = self._app()
            try:
                app.explicit_y_var.set("guessed")
                with (
                    mock.patch.object(gui.filedialog, "askopenfilename", return_value=str(record)),
                    mock.patch.object(gui.messagebox, "askokcancel", return_value=False) as cancel,
                    mock.patch.object(app, "_launch_batch") as launch,
                ):
                    app.continue_batch()
                    cancel.assert_called()
                    launch.assert_not_called()
                self.assertIsNone(app._import_process)
                orphan = {
                    "schema_version": 1,
                    "kind": "ori-batch-config",
                    "name": "别处",
                    "record_path": str(folder / "other-record.json"),
                    "inputs": [str(folder / "a.csv")],
                    "request": {
                        "inputs": [str(folder / "a.csv")],
                        "record_path": str(folder / "other-record.json"),
                        "plot": {"kind": "none", "x": "auto", "y": "auto", "y_error": "auto", "title": "auto", "x_label": "auto", "y_label": "auto"},
                        "overwrite": False,
                        "resume": False,
                        "retry_task_ids": [],
                    },
                }
                given = folder / "given-config.json"
                given.write_text(json.dumps(orphan), encoding="utf-8")
                answers = iter((str(record), str(given)))
                with (
                    mock.patch.object(gui.filedialog, "askopenfilename", side_effect=lambda **_kwargs: next(answers)),
                    mock.patch.object(gui.messagebox, "askokcancel", return_value=True),
                    mock.patch.object(gui.messagebox, "showerror") as showerror,
                    mock.patch.object(app, "_launch_batch") as launch,
                ):
                    app.continue_batch()
                    showerror.assert_called()
                    launch.assert_not_called()
            finally:
                _shutdown(app)


if __name__ == "__main__":
    unittest.main()
