"""GUI contracts and spawn tests, with opt-in real Tk integration; no Origin."""

from __future__ import annotations

import gc
import json
import multiprocessing
import os
import queue
import subprocess
import sys
import tempfile
import time
import types
import unittest
from pathlib import Path
from unittest import mock
import tkinter as tk

from openpyxl import Workbook, load_workbook

from origin_bridge import planning
from origin_bridge.gui import (
    _collect_supported,
    _default_output_for,
    _path_is_link,
    _plan_table_key,
    _read_options,
    _source_key,
    _validate_plan_save_target,
)
from origin_bridge.worker import execute_import_request, import_worker


class Value:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value


class YList:
    def __init__(self, values: list[str]):
        self.values = values

    def curselection(self):
        return tuple(range(len(self.values)))

    def get(self, index: int) -> str:
        return self.values[index]


class GenericGuiContractTests(unittest.TestCase):
    def test_read_options_default_to_empty_string_as_missing_value(self):
        options = _read_options("自动", "0", "自动识别")

        self.assertEqual(options["missing_values"], [""])
        self.assertEqual(options["header"], "auto")
        self.assertEqual(options["delimiter"], "auto")

    def test_empty_cells_do_not_turn_numeric_columns_into_text(self):
        with tempfile.TemporaryDirectory() as raw:
            source = Path(raw) / "values.csv"
            source.write_text("time,signal\n0,1\n1,\n2,3\n", encoding="utf-8")

            table = planning.inspect_inputs([source], options=_read_options("自动", "0", "自动识别"))["tables"][0]

        self.assertEqual([column["kind"] for column in table["columns"]], ["number", "number"])
        self.assertEqual(table["columns"][1]["missing"], 1)
        self.assertEqual(table["columns"][1]["sample"], ["1", None, "3"])

    def test_folder_expansion_uses_reader_extensions_and_natural_order(self):
        with tempfile.TemporaryDirectory() as raw:
            folder = Path(raw)
            first = folder / "sample10.ndjson"
            second = folder / "sample2.jsonl"
            ignored = folder / "notes.md"
            first.write_text('{"x":10}\n', encoding="utf-8")
            second.write_text('{"x":2}\n', encoding="utf-8")
            ignored.write_text("notes", encoding="utf-8")

            found, rejected = _collect_supported([folder, second])

        self.assertEqual([path.name for path in found], ["sample2.jsonl", "sample10.ndjson"])
        self.assertEqual(rejected, [])

    def test_explicit_missing_input_is_not_silently_ignored(self):
        with tempfile.TemporaryDirectory() as raw:
            with self.assertRaisesRegex(FileNotFoundError, "does not exist|不存在"):
                _collect_supported([Path(raw) / "missing.csv"])

    def test_rejected_warning_lists_original_reasons(self):
        from origin_bridge import gui
        from origin_bridge.gui import GeneralDataApp

        app = GeneralDataApp.__new__(GeneralDataApp)
        app.root = object()
        link = r"C:\data\dangling（无法跟随目录链接：目标不存在）"
        with mock.patch.object(gui.messagebox, "showwarning") as showwarning:
            app._apply_scan({"rejected": [r"C:\data\notes.md", link], "added": [], "announce": False})
        showwarning.assert_called_once()
        title, body = showwarning.call_args.args[:2]
        self.assertEqual(title, "部分来源未加入")
        self.assertNotIn("格式不支持", title)
        self.assertNotIn("格式不在支持列表", body)
        self.assertIn("notes.md", body)
        self.assertIn("无法跟随目录链接", body)
        self.assertIn(link, body)

    def test_default_workbook_path_does_not_overwrite_an_input_workbook(self):
        with tempfile.TemporaryDirectory() as raw:
            source = Path(raw) / "study.xlsx"
            source.touch()

            output = _default_output_for([source], "xlsx")

        self.assertNotEqual(output.resolve(), source.resolve())
        self.assertEqual(output.name, "study_imported.xlsx")

    def test_saved_plan_cannot_replace_a_source_or_planned_output(self):
        with tempfile.TemporaryDirectory() as raw:
            folder = Path(raw)
            plan = {
                "output": {"path": "result.xlsx"},
                "tables": [{"source": {"path": "source.json"}}],
            }
            with self.assertRaisesRegex(ValueError, "source.json"):
                _validate_plan_save_target(plan, folder / "source.json", folder)
            with self.assertRaisesRegex(ValueError, "result.xlsx"):
                _validate_plan_save_target(plan, folder / "result.xlsx", folder)

    def test_relative_plan_table_identity_uses_the_plan_directory(self):
        source = {"path": "data.csv", "options": {"sheet": None}}
        table = {"source": source}
        with tempfile.TemporaryDirectory() as raw:
            base = Path(raw)
            expected = str((base / "data.csv").resolve())
            self.assertEqual(_source_key(source, base)[0], expected.casefold())
            self.assertEqual(_plan_table_key(table, base), (expected.casefold(), None))

    def test_plot_error_mapping_uses_original_y_column_name_and_none_resets_axes(self):
        from origin_bridge.gui import GeneralDataApp

        app = GeneralDataApp.__new__(GeneralDataApp)
        app._current_key = ("source.csv", None)
        app._current_table = {
            "columns": [
                {"index": 0, "name": "Time"},
                {"index": 1, "name": "Signal"},
                {"index": 2, "name": "Uncertainty"},
            ]
        }
        app._plots = {}
        app.y_list = YList(["1: Signal (number)"])
        app.error_choice_var = Value("2: Uncertainty (number)")
        app.plot_kind_var = Value("line")
        app.x_choice_var = Value("0: Time (number)")
        app.title_var = Value("Signal")
        app.x_label_var = Value("Time")
        app.y_label_var = Value("Signal")

        app._save_current_plot()
        self.assertEqual(app._plots[app._current_key]["y_error"], {"Signal": 2})

        app.plot_kind_var = Value("none")
        app._save_current_plot()
        self.assertEqual(app._plots[app._current_key]["x"], None)
        self.assertEqual(app._plots[app._current_key]["y"], [])
        self.assertEqual(app._plots[app._current_key]["y_error"], {})

    def test_gui_plot_choices_build_a_plan_accepted_by_the_import_validator(self):
        from origin_bridge.gui import GeneralDataApp

        with tempfile.TemporaryDirectory() as raw:
            folder = Path(raw)
            source = folder / "signal.csv"
            source.write_text("Time,Signal,Error\n0,10,0.2\n1,12,0.3\n", encoding="utf-8")
            table = planning.inspect_inputs([source], options=_read_options("自动", "0", "自动识别"))["tables"][0]
            app = GeneralDataApp.__new__(GeneralDataApp)
            app._current_key = _source_key(table["source"])
            app._current_table = table
            app._plots = {}
            app._column_labels = {}
            app._loaded_plan = None
            app._plan_base_dir = None
            app._paths = [source]
            app._tables = [table]
            app.y_list = YList(["1: Signal (number)"])
            app.error_choice_var = Value("2: Error (number)")
            app.plot_kind_var = Value("line")
            app.x_choice_var = Value("0: Time (number)")
            app.title_var = Value("Signal over time")
            app.x_label_var = Value("Time")
            app.y_label_var = Value("Signal")
            app.header_var = Value("自动")
            app.skip_rows_var = Value("0")
            app.delimiter_var = Value("自动识别")
            app.keep_open_var = Value(False)

            plan = app._effective_plan(folder / "result.xlsx", "xlsx", False)
            prepared = planning.prepare_plan(plan)

        self.assertEqual(plan["tables"][0]["plot"]["y_error"], {"Signal": 2})
        self.assertEqual(prepared.tables[0].plot.y_error, ((1, 2),))

    def test_worker_prepares_then_exports_without_importing_tk(self):
        calls = []
        prepared = object()
        planning_module = types.ModuleType("origin_bridge.planning")
        planning_module.prepare_plan = lambda plan, base_dir=None: calls.append(("prepare", plan, base_dir)) or prepared
        exporter_module = types.ModuleType("origin_bridge.exporter")
        exporter_module.execute_import = lambda value: calls.append(("export", value)) or {
            "ok": True,
            "output": {"path": "result.xlsx", "format": "xlsx", "size_bytes": 10, "sha256": "a" * 64},
            "tables": [],
            "warnings": [],
        }
        plan = {"schema_version": 1}
        base_dir = Path("plan-folder")

        with mock.patch.dict(sys.modules, {
            "origin_bridge.planning": planning_module,
            "origin_bridge.exporter": exporter_module,
        }):
            result = execute_import_request(plan, base_dir)

        self.assertTrue(result["ok"])
        self.assertEqual(calls, [("prepare", plan, base_dir), ("export", prepared)])
        self.assertFalse(hasattr(sys.modules["origin_bridge.worker"], "tk"))
        json.dumps(result)

    def test_spawn_worker_exports_xlsx_without_opening_a_gui(self):
        with tempfile.TemporaryDirectory() as raw:
            folder = Path(raw)
            source = folder / "worker.csv"
            output = folder / "worker.xlsx"
            source.write_text("Time,Signal\n0,4\n1,5\n", encoding="utf-8")
            plan = planning.create_plan(
                [source], output, format="xlsx", options=_read_options("自动", "0", "自动识别")
            )
            context = multiprocessing.get_context("spawn")
            recv_conn, send_conn = context.Pipe(duplex=False)
            process = context.Process(target=import_worker, args=(send_conn, plan, None))
            process.start()
            send_conn.close()
            messages = []
            deadline = time.monotonic() + 30
            try:
                while time.monotonic() < deadline:
                    try:
                        if recv_conn.poll(0.1):
                            messages.append(recv_conn.recv())
                        elif not process.is_alive():
                            break
                    except (EOFError, OSError):
                        break
                process.join(timeout=2)
                self.assertFalse(process.is_alive(), "spawn worker did not finish")
            finally:
                if process.is_alive():
                    process.terminate()
                    process.join(timeout=2)
                recv_conn.close()
            exit_code = process.exitcode
            process.close()
            results = [message["result"] for message in messages if message.get("type") == "result"]
            self.assertEqual(exit_code, 0, messages)
            self.assertTrue(results and results[-1]["ok"], messages)
            self.assertTrue(output.is_file() and output.stat().st_size > 0)
            workbook = load_workbook(output, read_only=True, data_only=True)
            try:
                self.assertEqual(workbook["worker"].cell(3, 2).value, 5)
            finally:
                workbook.close()

    @unittest.skipUnless(os.environ.get("SPECTRA_TEST_GUI") == "1", "set SPECTRA_TEST_GUI=1 for the Tk and spawn-export integration check")
    def test_real_tk_remains_responsive_during_spawned_xlsx_export(self):
        from origin_bridge import gui
        from origin_bridge.gui import GeneralDataApp

        with tempfile.TemporaryDirectory() as raw:
            folder = Path(raw)
            source = folder / "signal.csv"
            source.write_text("Time,Signal\n0,10\n1,\n2,12\n", encoding="utf-8")
            plan_folder = folder / "saved-plans"
            plan_folder.mkdir()
            plan_path = plan_folder / "signal_plan.json"
            output = folder / "signal_result.xlsx"
            try:
                app = GeneralDataApp([source])
            except tk.TclError as exc:
                self.skipTest(f"Tk display is unavailable: {exc}")
            process = None
            try:
                app.root.withdraw()
                deadline = time.monotonic() + 15
                while app._busy and time.monotonic() < deadline:
                    app.root.update()
                    time.sleep(0.01)
                self.assertFalse(app._busy, "source inspection did not finish")
                self.assertEqual(app._tables[0]["columns"][1]["kind"], "number")
                self.assertEqual(app._tables[0]["columns"][1]["sample"], ["10", None, "12"])
                self.assertEqual(len(app.sample_tree.get_children()), 3)
                second_preview = app.sample_tree.item(app.sample_tree.get_children()[1], "values")
                self.assertEqual(second_preview[1], "")

                app.format_var.set("xlsx")
                app.plot_kind_var.set("none")
                with (
                    mock.patch.object(gui.filedialog, "asksaveasfilename", return_value=str(plan_path)),
                    mock.patch.object(gui.messagebox, "showerror") as showerror,
                ):
                    app.save_plan()
                    showerror.assert_not_called()
                saved_plan = json.loads(plan_path.read_text(encoding="utf-8"))
                saved_plan["tables"][0]["source"]["path"] = os.path.relpath(source, plan_folder)
                saved_plan["output"]["path"] = "future.xlsx"
                plan_path.write_text(json.dumps(saved_plan), encoding="utf-8")
                with (
                    mock.patch.object(gui.filedialog, "askopenfilename", return_value=str(plan_path)),
                    mock.patch.object(gui.messagebox, "showerror") as showerror,
                ):
                    app.load_plan()
                    deadline = time.monotonic() + 15
                    while app._busy and time.monotonic() < deadline:
                        app.root.update()
                        time.sleep(0.01)
                    self.assertFalse(app._busy, "relative-path plan inspection did not finish")
                    showerror.assert_not_called()
                self.assertEqual(app._paths, [source.resolve()])
                self.assertEqual(app._plan_base_dir, plan_folder.resolve())
                self.assertEqual(app.plot_kind_var.get(), "none")

                ticks = []

                def heartbeat():
                    ticks.append(time.monotonic())
                    if app._busy:
                        app.root.after(10, heartbeat)

                with (
                    mock.patch.object(gui.filedialog, "asksaveasfilename", return_value=str(output)),
                    mock.patch.object(gui.messagebox, "showinfo"),
                    mock.patch.object(gui.messagebox, "showerror") as showerror,
                    mock.patch.object(gui.messagebox, "showwarning"),
                ):
                    app.root.after(10, heartbeat)
                    app.export_data()
                    self.assertTrue(app._busy)
                    deadline = time.monotonic() + 45
                    while app._busy and time.monotonic() < deadline:
                        app.root.update()
                        time.sleep(0.01)
                    app.root.update()
                    self.assertFalse(app._busy, "spawned importer did not finish")
                    self.assertGreaterEqual(len(ticks), 2, "Tk callbacks stopped while the worker exported")
                    showerror.assert_not_called()
                self.assertEqual(str(app.export_button.cget("state")), "normal")
                self.assertTrue(output.is_file() and output.stat().st_size > 0)
                workbook = load_workbook(output, read_only=True, data_only=True)
                try:
                    sheet = workbook["signal"]
                    self.assertEqual(sheet.cell(3, 2).value, None)
                    self.assertEqual(sheet.cell(4, 2).value, 12)
                    self.assertIn("Provenance", workbook.sheetnames)
                finally:
                    workbook.close()
            finally:
                process = app._import_process
                if process is not None and process.is_alive():
                    process.terminate()
                    process.join(timeout=3)
                _shutdown_tk(app)
                app = None
                gc.collect()


class _Var:
    def __init__(self, value=""):
        self.value = value

    def set(self, value):
        self.value = value

    def get(self):
        return self.value


class _Pipe:
    def __init__(self, messages, raise_after=None, error=None):
        self.messages = list(messages)
        self.batch = 0
        self.max_batch = 0
        self.recv_count = 0
        self.raise_after = raise_after
        self.error = error or EOFError("pipe closed")

    def poll(self, _timeout=0):
        if self.raise_after is not None and self.recv_count >= self.raise_after:
            raise self.error
        return bool(self.messages)

    def recv(self):
        if self.raise_after is not None and self.recv_count >= self.raise_after:
            raise self.error
        if not self.messages:
            raise EOFError
        self.batch += 1
        self.recv_count += 1
        return self.messages.pop(0)

    def close(self):
        self.messages.clear()

    def note_batch(self):
        self.max_batch = max(self.max_batch, self.batch)
        self.batch = 0


class _Process:
    def __init__(self, alive=True, exit_code=0):
        self.alive = alive
        self.exitcode = exit_code
        self.closed = False

    def is_alive(self):
        return self.alive

    def join(self, timeout=None):
        self.alive = False

    def close(self):
        self.closed = True


def _contains_data_table(value):
    from origin_bridge.models import DataTable

    if isinstance(value, DataTable):
        return True
    if isinstance(value, dict):
        return any(_contains_data_table(item) for item in value.values())
    if isinstance(value, (list, tuple)):
        return any(_contains_data_table(item) for item in value)
    return False


class SummaryStoreTests(unittest.TestCase):
    def _options(self, header="自动", skip="0", delimiter="自动识别"):
        return _read_options(header, skip, delimiter)

    def test_incremental_reads_ignore_size_and_mtime_and_keep_going_after_a_bad_source(self):
        from origin_bridge import planning as planning_module
        from origin_bridge.source_summary import SummaryStore

        with tempfile.TemporaryDirectory() as raw:
            folder = Path(raw)
            first = folder / "first.csv"
            second = folder / "second.csv"
            changed = folder / "changed.csv"
            bad = folder / "bad.csv"
            wide = folder / "wide.csv"
            payload_a = b"a,b\n1,2\n"
            payload_b = b"a,b\n3,4\n"
            self.assertEqual(len(payload_a), len(payload_b))
            first.write_bytes(b"a,b\n1,2\n")
            second.write_bytes(b"a,b\n5,6\n")
            changed.write_bytes(payload_a)
            bad.write_bytes(b"")
            wide.write_text("x,y\n" + "\n".join(f"{index},{index}" for index in range(30)) + "\n", encoding="utf-8")
            stamp = 1_700_000_000
            os.utime(changed, (stamp, stamp))
            options = self._options()
            calls = []
            real_inspect = planning_module.inspect_inputs

            def spy(paths, options=None):
                calls.append([Path(path) for path in paths])
                return real_inspect(paths, options=options)

            store = SummaryStore()
            with mock.patch.object(planning_module, "inspect_inputs", side_effect=spy):
                first_result = store.inspect_path(first, options)
                second_result = store.inspect_path(second, options)
                bad_result = store.inspect_path(bad, options)
                wide_result = store.inspect_path(wide, options)
                changed_first = store.inspect_path(changed, options)
                misses = store.inspect_calls
                again = store.inspect_path(first, options)
                self.assertEqual(store.inspect_calls, misses)
                size_before = changed.stat().st_size
                mtime_before = changed.stat().st_mtime
                changed.write_bytes(payload_b)
                os.utime(changed, (stamp, stamp))
                self.assertEqual(changed.stat().st_size, size_before)
                self.assertEqual(changed.stat().st_mtime, mtime_before)
                rewritten = store.inspect_path(changed, options)
                reread = store.inspect_path(changed, options)
                shifted = store.inspect_path(first, self._options(header="无标题行"))

            self.assertTrue(all(len(batch) == 1 for batch in calls))
            self.assertTrue(first_result["ok"] and second_result["ok"] and wide_result["ok"] and changed_first["ok"])
            self.assertEqual(changed_first["tables"][0]["columns"][0]["sample"][0], "1")
            self.assertFalse(bad_result["ok"])
            self.assertIn("空", bad_result["error"])
            self.assertTrue(again["reused"])
            self.assertEqual(store.cache_hits, 2)
            self.assertFalse(rewritten["reused"])
            self.assertEqual(rewritten["tables"][0]["columns"][0]["sample"][0], "3")
            self.assertTrue(reread["reused"])
            self.assertEqual(store.inspect_calls, misses + 2)
            self.assertEqual([column["name"] for column in shifted["tables"][0]["columns"]], ["Column 1", "Column 2"])
            self.assertFalse(shifted["reused"])
            self.assertEqual(wide_result["tables"][0]["n_rows"], 30)
            self.assertLessEqual(len(wide_result["tables"][0]["columns"][0]["sample"]), 5)
            self.assertNotIn("values", wide_result["tables"][0]["columns"][0])
            self.assertFalse(_contains_data_table(wide_result))
            self.assertFalse(_contains_data_table(store._slots))


class ImportPollBudgetTests(unittest.TestCase):
    def _app(self):
        from origin_bridge.gui import GeneralDataApp

        app = GeneralDataApp.__new__(GeneralDataApp)
        app.status_var = _Var()
        app._busy = True
        app._import_result = None
        app._import_exit_waits = 0
        app._import_eof = False
        app._set_busy = lambda busy: setattr(app, "_busy", busy)
        app.root = types.SimpleNamespace(after=lambda *_args, **_kwargs: None)
        return app

    def test_progress_stream_is_bounded_per_poll_and_final_result_is_kept(self):
        from origin_bridge import gui

        app = self._app()
        messages = [{"type": "progress", "text": f"step {index}"} for index in range(100)]
        messages.append({"type": "result", "result": {"ok": True, "output": {"path": "out.xlsx", "size_bytes": 12}, "warnings": []}})
        pipe = _Pipe(messages)
        process = _Process(alive=True)
        app._import_recv = pipe
        app._import_process = process
        with mock.patch.object(gui.messagebox, "showinfo") as showinfo, mock.patch.object(gui.messagebox, "showerror") as showerror:
            while pipe.messages and process.alive:
                app._poll_import()
                pipe.note_batch()
                self.assertLessEqual(pipe.max_batch, 8)
            self.assertEqual(pipe.recv_count, 101)
            self.assertTrue(app._import_result["ok"])
            showinfo.assert_not_called()
            process.alive = False
            app._poll_import()
            showerror.assert_not_called()
            showinfo.assert_called_once()
        self.assertFalse(app._busy)
        self.assertIn("out.xlsx", app.status_var.get())

    def test_missing_result_after_process_exit_reports_the_exit_code(self):
        from origin_bridge import gui

        app = self._app()
        process = _Process(alive=False, exit_code=3)
        app._import_recv = _Pipe([])
        app._import_process = process
        with mock.patch.object(gui.messagebox, "showerror") as showerror, mock.patch.object(gui.messagebox, "showinfo") as showinfo:
            for _ in range(5):
                if not app._busy:
                    break
                app._poll_import()
            showinfo.assert_not_called()
            showerror.assert_called_once()
            self.assertIn("3", showerror.call_args.args[1])
            self.assertIn("退出码", showerror.call_args.args[1])
        self.assertFalse(app._busy)
        self.assertIn("退出码", app.status_var.get())
        self.assertTrue(process.closed)
        self.assertIsNone(app._import_process)
        self.assertIsNone(app._import_recv)

    def test_dead_worker_drains_a_long_backlog_and_keeps_one_result(self):
        from origin_bridge import gui

        app = self._app()
        messages = [{"type": "progress", "text": f"step {index}"} for index in range(20)]
        messages.append({"type": "result", "result": {"ok": True, "output": {"path": "dead.xlsx", "size_bytes": 4}, "warnings": []}})
        pipe = _Pipe(messages)
        process = _Process(alive=False, exit_code=0)
        app._import_recv = pipe
        app._import_process = process
        polls = 0
        with mock.patch.object(gui.messagebox, "showinfo") as showinfo, mock.patch.object(gui.messagebox, "showerror") as showerror:
            while app._busy and polls < 10:
                app._poll_import()
                self.assertLessEqual(pipe.batch, 8)
                pipe.note_batch()
                polls += 1
            showerror.assert_not_called()
            showinfo.assert_called_once()
            self.assertIn("dead.xlsx", showinfo.call_args.args[1])
        self.assertGreaterEqual(polls, 3)
        self.assertEqual(pipe.recv_count, 21)
        self.assertLessEqual(pipe.max_batch, 8)
        self.assertFalse(app._busy)
        self.assertEqual(app._import_result["output"]["path"], "dead.xlsx")
        self.assertIn("dead.xlsx", app.status_var.get())
        self.assertNotIn("step 0", app.status_var.get())
        self.assertTrue(process.closed)
        self.assertIsNone(app._import_recv)

    def test_mid_stream_pipe_errors_finish_instead_of_polling_forever(self):
        from origin_bridge import gui

        cases = (
            ("eof", EOFError("ended"), 9),
            ("oserror", OSError("broken pipe"), 4),
        )
        for name, error, exit_code in cases:
            with self.subTest(name=name):
                app = self._app()
                pipe = _Pipe(
                    [{"type": "progress", "text": "a"}, {"type": "progress", "text": "b"}, {"type": "progress", "text": "c"}],
                    raise_after=2,
                    error=error,
                )
                process = _Process(alive=False, exit_code=exit_code)
                app._import_recv = pipe
                app._import_process = process
                polls = 0
                with mock.patch.object(gui.messagebox, "showerror") as showerror, mock.patch.object(gui.messagebox, "showinfo") as showinfo:
                    while app._busy and polls < 6:
                        app._poll_import()
                        polls += 1
                    showinfo.assert_not_called()
                    showerror.assert_called_once()
                    self.assertIn(str(exit_code), showerror.call_args.args[1])
                self.assertLessEqual(polls, 3)
                self.assertFalse(app._busy)
                self.assertEqual(pipe.recv_count, 2)
                self.assertIn("退出码", app.status_var.get())
                self.assertTrue(process.closed)

    def test_non_dict_messages_are_skipped_and_the_later_result_is_kept(self):
        from origin_bridge import gui

        app = self._app()
        pipe = _Pipe([
            None,
            "nope",
            5,
            {"type": "progress", "text": "tail"},
            {"type": "result", "result": {"ok": True, "output": {"path": "kept.xlsx", "size_bytes": 1}, "warnings": ["watch"]}},
        ])
        process = _Process(alive=False, exit_code=0)
        app._import_recv = pipe
        app._import_process = process
        with mock.patch.object(gui.messagebox, "showinfo") as showinfo, mock.patch.object(gui.messagebox, "showerror") as showerror:
            while app._busy:
                app._poll_import()
                self.assertLessEqual(pipe.batch, 8)
                pipe.note_batch()
            showerror.assert_not_called()
            showinfo.assert_called_once()
            self.assertIn("watch", showinfo.call_args.args[1])
        self.assertEqual(pipe.recv_count, 5)
        self.assertEqual(app._import_result["output"]["path"], "kept.xlsx")
        self.assertIn("kept.xlsx", app.status_var.get())
        self.assertFalse(app._busy)


def _pump(app, seconds: float) -> None:
    deadline = time.monotonic() + seconds
    while app._busy and time.monotonic() < deadline:
        app.root.update()
        time.sleep(0.005)
    app.root.update()


def _shutdown_tk(app, thread=None) -> None:
    """Join the source thread and destroy the window. Caller then drops ``app``."""
    if thread is None:
        thread = getattr(app, "_thread", None)
    root = getattr(app, "root", None)
    try:
        if root is not None and root.winfo_exists():
            root.destroy()
    except tk.TclError:
        pass
    if thread is not None and getattr(thread, "is_alive", lambda: False)():
        thread.join(timeout=5)


class RealTkSourcePerformanceTests(unittest.TestCase):
    @unittest.skipUnless(os.environ.get("SPECTRA_TEST_GUI") == "1", "set SPECTRA_TEST_GUI=1 for the Tk and spawn-export integration check")
    def test_real_tk_reuses_summaries_and_keeps_going_after_one_bad_source(self):
        from origin_bridge import gui
        from origin_bridge.gui import GeneralDataApp

        with tempfile.TemporaryDirectory() as raw:
            folder = Path(raw)
            good = folder / "signal.csv"
            other = folder / "other.csv"
            extra = folder / "extra.csv"
            bad = folder / "bad.csv"
            payload_a = b"Time,Signal\n0,1\n"
            payload_b = b"Time,Signal\n9,8\n"
            self.assertEqual(len(payload_a), len(payload_b))
            good.write_bytes(payload_a)
            other.write_bytes(b"Time,Signal\n2,3\n")
            extra.write_bytes(b"Time,Signal\n4,5\n")
            os.utime(good, (1_700_000_000, 1_700_000_000))
            try:
                app = GeneralDataApp()
            except tk.TclError as exc:
                self.skipTest(f"Tk display is unavailable: {exc}")
            try:
                app.root.withdraw()
                with mock.patch.object(gui.messagebox, "showerror") as showerror:
                    app._add_paths([folder / "missing.csv"])
                    _pump(app, 10)
                    self.assertFalse(app._busy)
                    showerror.assert_called()
                self.assertEqual(app._paths, [])
                self.assertEqual(str(app.export_button.cget("state")), "normal")

                app._add_paths([good])
                _pump(app, 15)
                self.assertFalse(app._busy, "first source inspect did not finish")
                self.assertEqual(app._summaries.inspect_calls, 1)
                self.assertEqual(app._tables[0]["columns"][1]["sample"], ["1"])
                app.title_var.set("KEEP ME")
                app._save_current_plot()

                app._add_paths([other])
                _pump(app, 15)
                self.assertEqual(app._summaries.inspect_calls, 2)
                self.assertEqual(len(app._tables), 2)

                hits = app._summaries.cache_hits
                app.reload_inputs()
                _pump(app, 15)
                self.assertEqual(app._summaries.inspect_calls, 2)
                self.assertGreaterEqual(app._summaries.cache_hits, hits + 2)
                self.assertEqual(app.title_var.get(), "KEEP ME")

                mtime = good.stat().st_mtime
                size = good.stat().st_size
                good.write_bytes(payload_b)
                os.utime(good, (1_700_000_000, 1_700_000_000))
                self.assertEqual(good.stat().st_size, size)
                self.assertEqual(good.stat().st_mtime, mtime)
                app.reload_inputs()
                _pump(app, 15)
                self.assertEqual(app._summaries.inspect_calls, 3)
                self.assertEqual(app._tables[0]["columns"][1]["sample"], ["8"])
                self.assertNotEqual(app.title_var.get(), "KEEP ME")

                app.title_var.set("KEEP HEADER")
                app._save_current_plot()
                app.header_var.set("无标题行")
                app.reload_inputs()
                _pump(app, 15)
                self.assertGreaterEqual(app._summaries.inspect_calls, 5)
                names = [column["name"] for column in app._tables[0]["columns"]]
                self.assertEqual(names, ["Column 1", "Column 2"])
                self.assertNotEqual(app.title_var.get(), "KEEP HEADER")

                bad.write_bytes(b"")
                app.header_var.set("自动")
                app._add_paths([bad, extra])
                _pump(app, 15)
                self.assertFalse(app._busy)
                values = [app.table_tree.item(iid, "values") for iid in app.table_tree.get_children()]
                bad_rows = [row for row in values if Path(row[0]).name == "bad.csv"]
                extra_rows = [row for row in values if Path(row[0]).name == "extra.csv"]
                self.assertTrue(bad_rows and "错误" in bad_rows[0][2] and bad_rows[0][4])
                self.assertTrue(extra_rows and extra_rows[0][4] == "就绪")
                self.assertTrue(any(table["source"]["path"].endswith("extra.csv") for table in app._tables))
                self.assertFalse(any(table["source"]["path"].endswith("bad.csv") for table in app._tables))
                with self.assertRaisesRegex(ValueError, "不完整"):
                    app._effective_plan(folder / "blocked.xlsx", "xlsx", False)
                self.assertEqual(str(app.export_button.cget("state")), "normal")
                print(
                    "MEASURE incremental "
                    f"inspect_calls={app._summaries.inspect_calls} "
                    f"cache_hits={app._summaries.cache_hits} "
                    f"cache_misses={app._summaries.cache_misses}"
                )
            finally:
                _shutdown_tk(app)
                app = None
                gc.collect()

    @unittest.skipUnless(os.environ.get("SPECTRA_TEST_GUI") == "1", "set SPECTRA_TEST_GUI=1 for the Tk and spawn-export integration check")
    def test_real_tk_background_scan_and_chunked_refresh_stay_responsive_for_1000_sources(self):
        from origin_bridge import gui
        from origin_bridge.gui import GeneralDataApp
        from origin_bridge.readers import discover_files

        with tempfile.TemporaryDirectory() as raw:
            folder = Path(raw) / "sources"
            folder.mkdir()
            for index in range(1000):
                (folder / f"sample{index}.csv").write_text("x,y\n1,2\n", encoding="utf-8")
            (folder / "notes.md").write_text("ignore", encoding="utf-8")
            try:
                app = GeneralDataApp()
            except tk.TclError as exc:
                self.skipTest(f"Tk display is unavailable: {exc}")
            samples = []
            after_ids: list[str] = []

            def heartbeat():
                samples.append((time.monotonic(), len(app.table_tree.get_children())))
                if app._busy:
                    after_ids.append(app.root.after(10, heartbeat))

            try:
                app.root.withdraw()
                warmup = time.monotonic() + 0.5
                while time.monotonic() < warmup:
                    app.root.update()
                    time.sleep(0.01)
                after_ids.append(app.root.after(10, heartbeat))
                with (
                    mock.patch.object(gui.messagebox, "showwarning") as showwarning,
                    mock.patch.object(gui.messagebox, "showerror") as showerror,
                    mock.patch.object(gui.messagebox, "showinfo"),
                ):
                    app._add_paths([folder, folder / "notes.md"])
                    self.assertTrue(app._busy)
                    self.assertLess(len(app.table_tree.get_children()), 1000)
                    app._on_close()
                    self.assertTrue(app.root.winfo_exists())
                    deadline = time.monotonic() + 180
                    while app._busy and time.monotonic() < deadline:
                        app.root.update()
                        time.sleep(0.005)
                    app.root.update()
                    self.assertFalse(app._busy, "1000-source inspect did not finish")
                    showerror.assert_not_called()
                    showwarning.assert_called()
                names = [Path(app.table_tree.item(iid, "values")[0]).name for iid in app.table_tree.get_children()]
                self.assertEqual(names, [path.name for path in discover_files([folder])])
                self.assertEqual(len(names), 1000)
                self.assertEqual(app._summaries.inspect_calls, 1000)
                gaps = [samples[index][0] - samples[index - 1][0] for index in range(1, len(samples))]
                jumps = [samples[index][1] - samples[index - 1][1] for index in range(1, len(samples))]
                max_gap = max(gaps) if gaps else 999
                max_jump = max(jumps) if jumps else 1000
                partial = any(0 < count < 1000 for _stamp, count in samples)
                print(
                    "MEASURE tk1000 "
                    f"heartbeat_max_gap_s={max_gap:.4f} "
                    f"max_row_jump={max_jump} "
                    f"heartbeats={len(samples)} "
                    f"inspect_calls={app._summaries.inspect_calls} "
                    f"cache_hits={app._summaries.cache_hits}"
                )
                self.assertGreaterEqual(len(samples), 5)
                self.assertTrue(partial, "tree jumped to the full source list between heartbeats")
                self.assertLess(max_jump, 80, f"one heartbeat observed {max_jump} new rows")
                self.assertLess(max_gap, 0.25, f"Tk heartbeat stalled for {max_gap:.3f}s")
                self.assertEqual(str(app.export_button.cget("state")), "normal")
            finally:
                for after_id in after_ids:
                    try:
                        app.root.after_cancel(after_id)
                    except tk.TclError:
                        pass
                _shutdown_tk(app)
                app = None
                gc.collect()


class _DeadThread:
    def is_alive(self):
        return False

    def join(self, timeout=None):
        return None


def _release_links(path: Path) -> None:
    """Unlink junctions and symlinks without walking into their targets."""
    try:
        linked = _path_is_link(path)
    except OSError:
        return
    if linked:
        path.unlink()
        return
    if not path.is_dir():
        return
    for child in list(path.iterdir()):
        _release_links(child)


def _make_junction(link: Path, target: Path) -> None:
    proc = subprocess.run(["cmd", "/c", "mklink", "/J", str(link), str(target)], capture_output=True, text=True)
    if proc.returncode != 0:
        raise AssertionError((proc.stderr or proc.stdout or "mklink /J failed").strip())


class _LegacyDirEntry:
    """``os.DirEntry`` without ``is_junction``, as on Python 3.10 and 3.11."""

    def __init__(self, path: Path, *, directory: bool, reparse: bool):
        self.path = str(path)
        self.name = path.name
        self._directory = directory
        self._reparse = reparse

    def is_symlink(self) -> bool:
        return False

    def is_dir(self, *, follow_symlinks: bool = True) -> bool:
        return self._directory

    def is_file(self, *, follow_symlinks: bool = True) -> bool:
        return not self._directory

    def stat(self, *, follow_symlinks: bool = True):
        attributes = 0x10 if self._directory else 0x20
        if self._reparse:
            attributes |= 0x400
        return types.SimpleNamespace(st_file_attributes=attributes)


def _resolved_keys(paths) -> list[str]:
    return sorted(os.path.normcase(str(Path(path).resolve())) for path in paths)


class LinkDiscoveryTests(unittest.TestCase):
    @unittest.skipUnless(os.name == "nt", "junctions are a Windows directory link")
    def test_directory_links_are_followed_without_cycles_or_silent_drops(self):
        """Python 3.13 Windows: junctions are directories, symlinks need privilege.

        Acyclic junctions resolve to the same files as ``discover_files``.
        A junction cycle is deduped by the resolved ``seen_dirs`` set.
        ``discover_files`` / ``Path.rglob`` follows junctions and does not
        keep that set, so this test does not call it on the cycle. File and
        directory symlinks are followed when the process can create them;
        WinError 1314 skips only that creation.
        """
        from origin_bridge.readers import discover_files

        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            try:
                real = root / "real"
                real.mkdir()
                data = real / "data.csv"
                data.write_text("a,b\n1,2\n", encoding="utf-8")
                (real / "notes.md").write_text("skip", encoding="utf-8")
                scan = root / "scan"
                scan.mkdir()
                (scan / "local.csv").write_text("a,b\n5,6\n", encoding="utf-8")
                self.assertTrue(real.resolve().is_relative_to(root.resolve()))
                alias = scan / "alias"
                _make_junction(alias, real)
                self.assertFalse(alias.is_symlink())
                with os.scandir(scan) as entries:
                    entry = next(item for item in entries if item.name == "alias")
                    self.assertTrue(entry.is_dir(follow_symlinks=False))
                    self.assertFalse(entry.is_symlink())
                    if hasattr(alias, "is_junction"):
                        self.assertTrue(alias.is_junction())
                        self.assertTrue(entry.is_junction())

                found, rejected = _collect_supported([scan])
                cli = discover_files([scan])
                self.assertEqual(_resolved_keys(found), _resolved_keys(cli))
                self.assertEqual(sorted(path.name for path in found), ["data.csv", "local.csv"])
                self.assertEqual(rejected, [])

                loop = scan / "loop_j"
                _make_junction(loop, scan)
                found, rejected = _collect_supported([scan])
                self.assertEqual(sorted(path.name for path in found), ["data.csv", "local.csv"])
                self.assertEqual(len(found), 2)
                self.assertEqual(rejected, [])
                loop.unlink()
                self.assertTrue((scan / "local.csv").is_file())
                alias.unlink()
                self.assertTrue(data.is_file())

                gone = root / "gone"
                gone.mkdir()
                (gone / "lost.csv").write_text("a,b\n7,8\n", encoding="utf-8")
                dangling = scan / "dangling"
                _make_junction(dangling, gone)
                (gone / "lost.csv").unlink()
                gone.rmdir()
                found, rejected = _collect_supported([scan])
                self.assertEqual([path.name for path in found], ["local.csv"])
                self.assertTrue(any("dangling" in item and "目录链接" in item for item in rejected))
                dangling.unlink()

                file_error = None
                dir_error = None
                try:
                    os.symlink(data, scan / "via.csv")
                except OSError as exc:
                    file_error = exc
                try:
                    os.symlink(real, scan / "dirsym", target_is_directory=True)
                except OSError as exc:
                    dir_error = exc
                if file_error is None and dir_error is None:
                    found, rejected = _collect_supported([scan])
                    cli = discover_files([scan])
                    self.assertEqual(_resolved_keys(found), _resolved_keys(cli))
                    self.assertIn("data.csv", [path.name for path in found])
                    self.assertEqual(rejected, [])
                    print("LINK symlink=followed junction=followed cycle=canonical")
                else:
                    for label, exc in (("file symlink", file_error), ("dir symlink", dir_error)):
                        if exc is None:
                            continue
                        self.assertEqual(getattr(exc, "winerror", None), 1314, f"{label} failed differently: {exc}")
                    print(
                        "LINK junction=followed cycle=canonical discover_files=acyclic match "
                        "symlink_skip=WinError 1314 "
                        f"file={file_error} dir={dir_error}"
                    )
            finally:
                _release_links(root)

    @unittest.skipUnless(os.name == "nt", "junctions are a Windows directory link")
    def test_legacy_reparse_files_stay_files_without_is_junction(self):
        """Python 3.10/3.11 DirEntry has no ``is_junction``. This process does not run those interpreters.

        A reparse file must be collected. A reparse directory is walked once.
        A broken directory link is reported as a directory link.
        """
        from origin_bridge import gui

        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            scan = root / "scan"
            scan.mkdir()
            recall = scan / "recall.csv"
            recall.write_text("a,b\n1,2\n", encoding="utf-8")
            local = scan / "local.csv"
            local.write_text("a,b\n3,4\n", encoding="utf-8")
            nested = scan / "nested"
            nested.mkdir()
            inner = nested / "inner.csv"
            inner.write_text("a,b\n5,6\n", encoding="utf-8")
            alias = scan / "nested_link"
            _make_junction(alias, nested)
            gone = root / "gone"
            gone.mkdir()
            dangling = scan / "dangling"
            _make_junction(dangling, gone)
            gone.rmdir()
            entries = [
                _LegacyDirEntry(recall, directory=False, reparse=True),
                _LegacyDirEntry(local, directory=False, reparse=False),
                _LegacyDirEntry(nested, directory=True, reparse=True),
                _LegacyDirEntry(alias, directory=True, reparse=True),
                _LegacyDirEntry(scan, directory=True, reparse=True),
                _LegacyDirEntry(dangling, directory=True, reparse=True),
            ]
            self.assertFalse(hasattr(entries[0], "is_junction"))
            root_key = os.path.normcase(str(scan.resolve()))
            real_scandir = gui.os.scandir

            def scanning(path):
                key = os.path.normcase(str(Path(path).resolve(strict=False)))
                if key == root_key:
                    return iter(entries)
                return real_scandir(path)

            try:
                with mock.patch.object(gui.os, "scandir", side_effect=scanning):
                    found, rejected = _collect_supported([scan])
            finally:
                _release_links(root)

        names = [path.name for path in found]
        self.assertEqual(sorted(names), ["inner.csv", "local.csv", "recall.csv"])
        self.assertEqual(names.count("inner.csv"), 1)
        self.assertEqual(len(rejected), 1)
        self.assertIn("dangling", rejected[0])
        self.assertIn("目录链接", rejected[0])
        self.assertNotIn("recall.csv", rejected[0])
        print(
            "LINK legacy_direntry=no_is_junction "
            f"python={sys.version.split()[0]} "
            "reparse_file=collected reparse_dir=traversed cycle=deduped broken=目录链接 "
            "interpreter_3_10=not_executed interpreter_3_11=not_executed"
        )


class SnapshotContractTests(unittest.TestCase):
    def test_batch_source_snapshot_lists_errors_without_full_rows(self):
        from origin_bridge.gui import GeneralDataApp
        from origin_bridge.source_summary import SummaryStore

        with tempfile.TemporaryDirectory() as raw:
            wide = Path(raw) / "wide.csv"
            wide.write_text("x,y\n" + "\n".join(f"{index},{index}" for index in range(30)) + "\n", encoding="utf-8")
            store = SummaryStore()
            outcome = store.inspect_path(wide, _read_options("自动", "0", "自动识别"))
            app = GeneralDataApp.__new__(GeneralDataApp)
            app._paths = [wide]
            app._tables = list(outcome["tables"]) + [{
                "error": "stale",
                "source": {"path": str(wide)},
                "columns": [{"values": list(range(100))}],
            }]
            app._source_errors = {"bad-key": "空文件"}
            app._summaries = store
            snap = GeneralDataApp.batch_source_snapshot(app)

        self.assertEqual(set(snap), {"paths", "tables", "errors", "inspect_calls", "cache_hits"})
        self.assertEqual(snap["errors"], {"bad-key": "空文件"})
        self.assertEqual(len(snap["tables"]), 1)
        self.assertEqual(snap["tables"][0]["n_rows"], 30)
        column = snap["tables"][0]["columns"][0]
        self.assertLessEqual(len(column["sample"]), 5)
        self.assertNotIn("values", column)
        self.assertNotIn("stale", json.dumps(snap["tables"]))
        self.assertFalse(_contains_data_table(snap))
        self.assertEqual(snap["inspect_calls"], 1)


class SourceJobFailureTests(unittest.TestCase):
    def test_unexpected_thread_death_confirms_once_then_reports_the_fallback(self):
        from collections import deque

        from origin_bridge import gui
        from origin_bridge.gui import GeneralDataApp

        app = GeneralDataApp.__new__(GeneralDataApp)
        app._busy = True
        app._thread = _DeadThread()
        app._thread_kind = "scan"
        app._thread_queue = queue.Queue()
        app._pending_outcomes = deque()
        app._job_terminal = False
        app._job_error = None
        app._done_message = None
        app._poll_confirms = 0
        app._paths = []
        app._tables = []
        app._source_errors = {}
        app.status_var = _Var()
        app.root = types.SimpleNamespace(after=lambda *_args, **_kwargs: None)
        app._ensure_selection = lambda: None
        app._set_busy = lambda busy: setattr(app, "_busy", busy)
        with mock.patch.object(gui.messagebox, "showerror") as showerror:
            app._poll_thread()
            showerror.assert_not_called()
            self.assertTrue(app._busy)
            self.assertEqual(app._poll_confirms, 1)
            self.assertIsNotNone(app._thread)
            app._poll_thread()
            showerror.assert_called_once()
            self.assertIn("后台读取意外结束", showerror.call_args.args[1])
            self.assertNotIn("读取来源失败", showerror.call_args.args[1])
        self.assertFalse(app._busy)
        self.assertIsNone(app._thread)
        self.assertEqual(app.status_var.get(), "读取失败。")


def _tree_rows(app):
    return [(iid, app.table_tree.item(iid, "values")) for iid in app.table_tree.get_children()]


def _select_path_sheet(app, name: str, sheet: str):
    for iid, values in _tree_rows(app):
        if Path(values[0]).name == name and values[1] == sheet:
            app.table_tree.selection_set(iid)
            app.table_tree.focus(iid)
            app._on_table_selection()
            return iid
    raise AssertionError(f"missing row {name} {sheet!r}")


def _column_names(app) -> list[str]:
    return [app.column_tree.item(iid, "values")[1] for iid in app.column_tree.get_children()]


class RealTkSourceRepairTests(unittest.TestCase):
    @unittest.skipUnless(os.environ.get("SPECTRA_TEST_GUI") == "1", "set SPECTRA_TEST_GUI=1 for the Tk and spawn-export integration check")
    def test_scan_exception_restores_controls_and_blocks_an_unfinished_source(self):
        from origin_bridge import gui
        from origin_bridge.gui import GeneralDataApp

        with tempfile.TemporaryDirectory() as raw:
            folder = Path(raw)
            good = folder / "signal.csv"
            good.write_text("Time,Signal\n0,1\n", encoding="utf-8")
            try:
                app = GeneralDataApp()
            except tk.TclError as exc:
                self.skipTest(f"Tk display is unavailable: {exc}")
            try:
                app.root.withdraw()
                with (
                    mock.patch.object(gui.messagebox, "showerror") as showerror,
                    mock.patch.object(app, "_inspect_paths_worker", side_effect=RuntimeError("inspect boom")),
                ):
                    app._add_paths([good])
                    _pump(app, 10)
                self.assertFalse(app._busy)
                self.assertIsNone(app._thread)
                self.assertEqual(str(app.export_button.cget("state")), "normal")
                showerror.assert_called()
                self.assertIn("inspect boom", showerror.call_args.args[1])
                self.assertIn("读取来源失败", showerror.call_args.args[1])
                self.assertNotIn("意外结束", showerror.call_args.args[1])
                self.assertEqual(len(app._paths), 1)
                with self.assertRaisesRegex(ValueError, "尚未完成"):
                    app._effective_plan(folder / "blocked.xlsx", "xlsx", False)
            finally:
                _shutdown_tk(app)
                app = None
                gc.collect()

    @unittest.skipUnless(os.environ.get("SPECTRA_TEST_GUI") == "1", "set SPECTRA_TEST_GUI=1 for the Tk and spawn-export integration check")
    def test_destroy_during_scan_does_not_schedule_after_on_a_dead_window(self):
        import threading

        from origin_bridge.gui import GeneralDataApp

        with tempfile.TemporaryDirectory() as raw:
            source = Path(raw) / "signal.csv"
            source.write_text("Time,Signal\n0,1\n", encoding="utf-8")
            try:
                app = GeneralDataApp()
            except tk.TclError as exc:
                self.skipTest(f"Tk display is unavailable: {exc}")
            started = threading.Event()

            def block(_paths, _options, _update):
                started.set()
                time.sleep(1.2)

            errors = []
            thread = None
            try:
                app.root.withdraw()
                app.root.report_callback_exception = lambda *args: errors.append(args)
                with mock.patch.object(app, "_inspect_paths_worker", block):
                    app._add_paths([source])
                    self.assertTrue(started.wait(5))
                    self.assertTrue(app._busy)
                    for _ in range(8):
                        app.root.update()
                        time.sleep(0.02)
                    thread = app._thread
                    app.root.destroy()
                self.assertEqual(errors, [])
            finally:
                _shutdown_tk(app, thread)
                app = None
                gc.collect()

    @unittest.skipUnless(os.environ.get("SPECTRA_TEST_GUI") == "1", "set SPECTRA_TEST_GUI=1 for the Tk and spawn-export integration check")
    def test_direct_destroy_closes_export_pipe_and_reaps_after_natural_exit(self):
        from origin_bridge import gui
        from origin_bridge.gui import GeneralDataApp

        class _LiveExport:
            def __init__(self):
                self._alive = True
                self.exitcode = None
                self.closed = False
                self.terminated = False
                self.join_timeouts: list[float | None] = []

            def is_alive(self):
                if self.closed:
                    raise ValueError("process object is closed")
                return self._alive

            def join(self, timeout=None):
                if self.closed:
                    raise ValueError("process object is closed")
                self.join_timeouts.append(timeout)
                if self._alive and timeout is None:
                    raise AssertionError("export reaper joined without a bound")
                if self._alive and timeout:
                    time.sleep(timeout)

            def close(self):
                if self._alive:
                    raise ValueError("Cannot close a process while it is still running")
                self.closed = True

            def terminate(self):
                self.terminated = True

            def kill(self):
                self.terminated = True

            def finish(self):
                self._alive = False
                self.exitcode = 0

        try:
            app = GeneralDataApp()
        except tk.TclError as exc:
            self.skipTest(f"Tk display is unavailable: {exc}")
        ctx = multiprocessing.get_context("spawn")
        recv, send = ctx.Pipe(duplex=False)
        send.close()
        process = _LiveExport()
        try:
            app.root.withdraw()
            with mock.patch.object(gui.messagebox, "showwarning") as showwarning:
                app._busy = True
                app._on_close()
                self.assertTrue(app.root.winfo_exists())
                showwarning.assert_called_once()
                self.assertIn("任务进行中", showwarning.call_args.args[0])
            fired = []
            app._import_process = process
            app._import_recv = recv
            app._after(30, lambda: fired.append("poll"))
            scheduled = app._poll_after
            self.assertIsNotNone(scheduled)
            cancelled = []
            real_cancel = app.root.after_cancel

            def spy_cancel(after_id, *args, **kwargs):
                cancelled.append(after_id)
                return real_cancel(after_id, *args, **kwargs)

            app.root.after_cancel = spy_cancel
            app.root.destroy()
            self.assertIsNone(app._poll_after)
            self.assertIsNone(app._hook_after)
            self.assertIn(scheduled, cancelled)
            self.assertEqual(fired, [])
            self.assertTrue(recv.closed)
            self.assertIsNone(app._import_recv)
            self.assertIsNone(app._import_process)
            self.assertTrue(process.is_alive())
            self.assertFalse(process.closed)
            self.assertFalse(process.terminated)
            deadline = time.monotonic() + 2
            while not process.join_timeouts and time.monotonic() < deadline:
                time.sleep(0.01)
            self.assertTrue(process.join_timeouts)
            self.assertTrue(all(timeout is not None and timeout <= 0.2 for timeout in process.join_timeouts))
            process.finish()
            deadline = time.monotonic() + 2
            while not process.closed and time.monotonic() < deadline:
                time.sleep(0.01)
            self.assertTrue(process.closed)
            self.assertEqual(process.exitcode, 0)
            self.assertFalse(process.terminated)
            self.assertEqual(fired, [])
            print(
                "DESTROY direct=export after_cleared=1 pipe_closed=1 "
                f"reaper_join_timeouts={process.join_timeouts} process_closed_after_exit=1 terminated=0"
            )
        finally:
            process.finish()
            deadline = time.monotonic() + 2
            while not process.closed and time.monotonic() < deadline:
                time.sleep(0.01)
            try:
                recv.close()
            except OSError:
                pass
            _shutdown_tk(app)
            app = None
            gc.collect()

    @unittest.skipUnless(os.environ.get("SPECTRA_TEST_GUI") == "1", "set SPECTRA_TEST_GUI=1 for the Tk and spawn-export integration check")
    def test_error_row_clears_controls_and_removed_plots_do_not_return(self):
        from origin_bridge.gui import GeneralDataApp

        with tempfile.TemporaryDirectory() as raw:
            folder = Path(raw)
            good = folder / "signal.csv"
            bad = folder / "bad.csv"
            good.write_text("Time,Signal\n0,1\n", encoding="utf-8")
            bad.write_bytes(b"")
            try:
                app = GeneralDataApp()
            except tk.TclError as exc:
                self.skipTest(f"Tk display is unavailable: {exc}")
            try:
                app.root.withdraw()
                app._add_paths([good])
                _pump(app, 15)
                app.title_var.set("KEEP TITLE")
                app.x_label_var.set("KEEP X")
                app.y_label_var.set("KEEP Y")
                app.plot_kind_var.set("scatter")
                app._save_current_plot()
                good_key = app._table_identity(app._tables[0]["source"])
                app._column_labels[good_key] = {"Signal": "OLD LABEL"}
                app._add_paths([bad])
                _pump(app, 15)
                _select_path_sheet(app, "bad.csv", "")
                self.assertEqual(app.title_var.get(), "")
                self.assertEqual(app.x_label_var.get(), "")
                self.assertEqual(app.y_label_var.get(), "")
                self.assertEqual(app.x_choice_var.get(), "行号")
                self.assertEqual(app.error_choice_var.get(), "无")
                self.assertEqual(app.plot_kind_var.get(), "none")
                self.assertEqual(app.y_list.size(), 0)
                self.assertIn("来源读取失败", app.status_var.get())
                bad_row = next(values for _iid, values in _tree_rows(app) if Path(values[0]).name == "bad.csv")
                self.assertEqual(bad_row[2], "错误")
                self.assertTrue(bad_row[4])
                with self.assertRaisesRegex(ValueError, "不完整"):
                    app._effective_plan(folder / "blocked.xlsx", "xlsx", False)

                _select_path_sheet(app, "signal.csv", "")
                self.assertEqual(app.title_var.get(), "KEEP TITLE")
                self.assertEqual(app.x_label_var.get(), "KEEP X")
                self.assertEqual(app.y_label_var.get(), "KEEP Y")
                self.assertEqual(app.plot_kind_var.get(), "scatter")
                self.assertGreater(app.y_list.size(), 0)

                bad.write_text("p,q,r\n1,2,3\n", encoding="utf-8")
                app.reload_inputs()
                _pump(app, 15)
                self.assertEqual(app._source_errors, {})
                self.assertFalse(any(values[2] == "错误" for _iid, values in _tree_rows(app)))
                _select_path_sheet(app, "bad.csv", "")
                self.assertEqual(_column_names(app), ["p", "q", "r"])
                self.assertNotEqual(app.title_var.get(), "KEEP TITLE")
                _select_path_sheet(app, "signal.csv", "")
                self.assertEqual(app.title_var.get(), "KEEP TITLE")

                _select_path_sheet(app, "signal.csv", "")
                app.remove_selected_sources()
                self.assertFalse(any(Path(values[0]).name == "signal.csv" for _iid, values in _tree_rows(app)))
                good.write_text("a,b,c\n9,8,7\n", encoding="utf-8")
                app._add_paths([good])
                _pump(app, 15)
                _select_path_sheet(app, "signal.csv", "")
                self.assertEqual(_column_names(app), ["a", "b", "c"])
                self.assertNotEqual(app.title_var.get(), "KEEP TITLE")
                remembered = repr(app._column_labels) + repr(app._plots)
                self.assertNotIn("OLD LABEL", remembered)
                self.assertNotIn("KEEP X", remembered)
            finally:
                _shutdown_tk(app)
                app = None
                gc.collect()

    @unittest.skipUnless(os.environ.get("SPECTRA_TEST_GUI") == "1", "set SPECTRA_TEST_GUI=1 for the Tk and spawn-export integration check")
    def test_excel_sheets_keep_their_own_plots_through_plan_load_and_reload(self):
        from origin_bridge import gui
        from origin_bridge.gui import GeneralDataApp

        with tempfile.TemporaryDirectory() as raw:
            folder = Path(raw)
            workbook_path = folder / "book.xlsx"
            workbook = Workbook()
            alpha = workbook.active
            alpha.title = "alpha"
            alpha.append(["x", "y"])
            alpha.append([1, 2])
            beta = workbook.create_sheet("beta")
            beta.append(["x", "y"])
            beta.append([3, 4])
            beta.append([5, 6])
            gamma = workbook.create_sheet("gamma")
            gamma.append(["p", "q"])
            gamma.append([7, 8])
            workbook.save(workbook_path)
            workbook.close()
            companion = folder / "companion.csv"
            companion.write_text("Time,Signal\n0,1\n", encoding="utf-8")
            try:
                app = GeneralDataApp()
            except tk.TclError as exc:
                self.skipTest(f"Tk display is unavailable: {exc}")
            try:
                app.root.withdraw()
                app._add_paths([companion, workbook_path])
                _pump(app, 20)
                sheets = [values[1] for _iid, values in _tree_rows(app) if Path(values[0]).name == "book.xlsx"]
                self.assertEqual(sheets, ["alpha", "beta", "gamma"])
                counts = [values[2] for _iid, values in _tree_rows(app) if Path(values[0]).name == "book.xlsx"]
                self.assertEqual(counts, ["1", "2", "1"])
                self.assertEqual(len(app._summaries._slots), 2)
                _select_path_sheet(app, "companion.csv", "")
                app.title_var.set("CSV TITLE")
                app._save_current_plot()
                for sheet, title in (("alpha", "ALPHA TITLE"), ("beta", "BETA TITLE"), ("gamma", "GAMMA TITLE")):
                    _select_path_sheet(app, "book.xlsx", sheet)
                    app.title_var.set(title)
                    app._save_current_plot()
                _select_path_sheet(app, "book.xlsx", "alpha")
                self.assertEqual(app.title_var.get(), "ALPHA TITLE")
                _select_path_sheet(app, "book.xlsx", "gamma")
                self.assertEqual(_column_names(app), ["p", "q"])
                self.assertEqual(app.title_var.get(), "GAMMA TITLE")
                _select_path_sheet(app, "book.xlsx", "beta")
                self.assertEqual(app.title_var.get(), "BETA TITLE")

                _select_path_sheet(app, "book.xlsx", "alpha")
                app.remove_selected_sources()
                _select_path_sheet(app, "companion.csv", "")
                self.assertEqual(app.title_var.get(), "CSV TITLE")
                self.assertFalse(any(Path(values[0]).name == "book.xlsx" for _iid, values in _tree_rows(app)))
                app._add_paths([workbook_path])
                _pump(app, 20)
                _select_path_sheet(app, "book.xlsx", "alpha")
                self.assertNotEqual(app.title_var.get(), "ALPHA TITLE")
                _select_path_sheet(app, "companion.csv", "")
                self.assertEqual(app.title_var.get(), "CSV TITLE")

                by_sheet = {table["source"].get("sheet"): table for table in app._tables if str(table["source"].get("path", "")).endswith("book.xlsx")}
                calls_before = app._summaries.inspect_calls
                slots_before = len(app._summaries._slots)

                def plan_for(order, titles):
                    tables = []
                    for sheet in order:
                        table = by_sheet[sheet]
                        source = table["source"]
                        tables.append({
                            "source": {field: source[field] for field in ("path", "sha256", "options")},
                            "name": table["name"],
                            "plot": {
                                "kind": "line",
                                "x": 0,
                                "y": [1],
                                "y_error": {},
                                "title": titles[sheet],
                                "x_label": f"X-{sheet}",
                                "y_label": f"Y-{sheet}",
                            },
                            "column_labels": {"y": f"label-{sheet}"},
                        })
                    return {
                        "schema_version": 1,
                        "output": {"path": str(folder / "out.xlsx"), "format": "xlsx", "overwrite": False, "keep_open": False},
                        "tables": tables,
                    }

                reversed_plan = plan_for(("gamma", "alpha", "beta"), {"gamma": "G-PLAN", "alpha": "A-PLAN", "beta": "B-PLAN"})
                plan_path = folder / "sheets.json"
                plan_path.write_text(json.dumps(reversed_plan), encoding="utf-8")
                with mock.patch.object(gui.filedialog, "askopenfilename", return_value=str(plan_path)), mock.patch.object(gui.messagebox, "showerror") as showerror:
                    app.load_plan()
                    _pump(app, 20)
                    showerror.assert_not_called()
                loaded_sheets = [values[1] for _iid, values in _tree_rows(app)]
                self.assertEqual(loaded_sheets, ["gamma", "alpha", "beta"])
                # The first inspect requested header "auto". Saved sheet options
                # store the resolved header boolean, so this reload's fingerprint
                # misses once for the workbook. All three sheets share that read.
                self.assertEqual(app._summaries.inspect_calls, calls_before + 1)
                self.assertLessEqual(len(app._summaries._slots), slots_before)
                for sheet, title in (("gamma", "G-PLAN"), ("alpha", "A-PLAN"), ("beta", "B-PLAN")):
                    _select_path_sheet(app, "book.xlsx", sheet)
                    self.assertEqual(app.title_var.get(), title)
                    self.assertEqual(app.x_label_var.get(), f"X-{sheet}")
                self.assertEqual(app._loaded_plan["schema_version"], 1)

                single = plan_for(("beta",), {"beta": "BETA ONLY"})
                single["tables"][0]["column_labels"] = {"y": "BETA LABEL"}
                single_path = folder / "beta.json"
                single_path.write_text(json.dumps(single), encoding="utf-8")
                with mock.patch.object(gui.filedialog, "askopenfilename", return_value=str(single_path)), mock.patch.object(gui.messagebox, "showerror") as showerror:
                    app.load_plan()
                    _pump(app, 20)
                    showerror.assert_not_called()
                self.assertEqual([values[1] for _iid, values in _tree_rows(app)], ["beta"])
                self.assertEqual(app.title_var.get(), "BETA ONLY")
                self.assertIsNotNone(app._loaded_plan)
                app.reload_inputs()
                _pump(app, 20)
                self.assertIsNone(app._loaded_plan)
                self.assertEqual([values[1] for _iid, values in _tree_rows(app)], ["alpha", "beta", "gamma"])
                for sheet in ("alpha", "beta", "gamma"):
                    _select_path_sheet(app, "book.xlsx", sheet)
                    self.assertNotEqual(app.title_var.get(), "BETA ONLY")
                    self.assertNotEqual(app.x_label_var.get(), "X-beta")
                remembered = repr(app._column_labels) + repr(app._plots)
                self.assertNotIn("BETA LABEL", remembered)
                self.assertNotIn("BETA ONLY", remembered)
                print(
                    "MEASURE excel "
                    f"inspect_calls={app._summaries.inspect_calls} "
                    f"slots={len(app._summaries._slots)} "
                    f"reload_sheets={[values[1] for _iid, values in _tree_rows(app)]}"
                )
            finally:
                _shutdown_tk(app)
                app = None
                gc.collect()


if __name__ == "__main__":
    unittest.main()
