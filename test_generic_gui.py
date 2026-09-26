"""GUI contracts and spawn tests, with opt-in real Tk integration; no Origin."""

from __future__ import annotations

import json
import multiprocessing
import os
import sys
import tempfile
import time
import types
import unittest
from pathlib import Path
from unittest import mock
import tkinter as tk

from openpyxl import load_workbook

from origin_bridge import planning
from origin_bridge.gui import (
    _collect_supported,
    _default_output_for,
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
                app.root.destroy()


if __name__ == "__main__":
    unittest.main()
