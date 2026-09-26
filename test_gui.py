"""GUI status, cache, and spawn-worker checks without opening Origin or Tk."""

from __future__ import annotations

import multiprocessing
import os
import tempfile
import time
import unittest
import zipfile
from pathlib import Path
from unittest import mock

import tkinter as tk

import spectra_to_origin as sto
from gui_export import execute_export_request, export_worker


def _spectrum(name: str, xs: tuple[str, ...], ys: tuple[str, ...]) -> sto.Spectrum:
    return sto.Spectrum(
        path=Path(name),
        x_text=xs,
        y_text=ys,
        long_name=Path(name).stem,
        comment="",
        temperature_tag=None,
    )


class GuiStatusTests(unittest.TestCase):
    def test_auto_status_reports_layout_and_column_count_for_each_group(self) -> None:
        shared = ("0", "1", "2")
        spectra = [
            _spectrum("a_1.txt", shared, ("1", "2", "3")),
            _spectrum("a_2.txt", shared, ("3", "2", "1")),
            _spectrum("b_1.txt", ("0", "2", "4"), ("1", "2", "3")),
            _spectrum("b_2.txt", ("0", "3", "6"), ("3", "2", "1")),
        ]

        text = sto._status_text(spectra, "auto", ["A", "A", "B", "B"])

        self.assertIn("A: XYYY, 3 列", text)
        self.assertIn("B: XYXY, 4 列", text)

    def test_fixed_xyyy_status_identifies_group_with_mismatched_x(self) -> None:
        spectra = [
            _spectrum("a.txt", ("0", "1"), ("1", "2")),
            _spectrum("b.txt", ("0", "2"), ("3", "4")),
        ]

        text = sto._status_text(spectra, "XYYY", ["one", "one"])

        self.assertIn("one: XYYY", text)
        self.assertIn("X 网格不一致", text)

    def test_same_filename_in_different_folders_remains_distinguishable(self) -> None:
        paths = [Path("C:/first/sample.txt"), Path("D:/second/sample.txt")]

        labels = sto.SpectraToOriginApp._list_labels(paths, ["A", "A"])

        self.assertNotEqual(labels[0], labels[1])
        self.assertIn("first", labels[0])
        self.assertIn("second", labels[1])


class ParsedCacheTests(unittest.TestCase):
    def test_cache_reloads_when_size_or_mtime_changes(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            path = Path(raw).resolve() / "sample.txt"
            path.write_text("0 1\n1 2\n", encoding="utf-8")
            app = sto.SpectraToOriginApp.__new__(sto.SpectraToOriginApp)
            app.files = [path]
            app._cached_key = None
            app._cached_spectra = []

            first = app._parsed_spectra()
            before = path.stat()
            path.write_text("0 9\n1 2\n", encoding="utf-8")
            os.utime(path, ns=(before.st_atime_ns, before.st_mtime_ns + 1_000_000))
            second = app._parsed_spectra()

            self.assertEqual(first[0].y_text[0], "1")
            self.assertEqual(second[0].y_text[0], "9")


class ExportWorkerTests(unittest.TestCase):
    def test_opju_worker_refuses_to_attach_to_running_origin(self) -> None:
        request = {
            "kind": "opju",
            "files": [],
            "groups": [],
            "output": "unused.opju",
            "layout": "auto",
            "axis_labels": ("2theta", "deg", "Intensity", "a.u."),
            "show_origin": False,
            "keep_open": False,
            "also_xlsx": False,
        }
        with mock.patch.object(sto, "origin_process_running", return_value=True), mock.patch.object(
            sto, "export_origin_project"
        ) as export:
            result = execute_export_request(request)

        self.assertFalse(result["ok"])
        self.assertIn("请先保存并关闭 Origin", result["error"])
        export.assert_not_called()

    def test_xlsx_export_runs_in_a_spawned_worker(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            folder = Path(raw).resolve()
            source = folder / "sample.txt"
            source.write_text("0 1\n1 2\n2 3\n", encoding="utf-8")
            spectrum = sto.parse_spectrum(source)
            request = {
                "kind": "xlsx",
                "files": [source],
                "groups": [("all", [spectrum])],
                "output": str(folder / "out.xlsx"),
                "layout": "auto",
                "axis_labels": ("q", "nm^-1", "Counts", "a.u."),
                "show_origin": False,
                "keep_open": False,
                "also_xlsx": False,
            }
            context = multiprocessing.get_context("spawn")
            recv_conn, send_conn = context.Pipe(duplex=False)
            process = context.Process(target=export_worker, args=(send_conn, request))
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
                process.join(timeout=1)
                self.assertFalse(process.is_alive(), "spawn worker did not finish")
            finally:
                if process.is_alive():
                    process.terminate()
                    process.join(timeout=2)
                recv_conn.close()

            results = [message["result"] for message in messages if message["type"] == "result"]
            self.assertEqual(process.exitcode, 0)
            self.assertTrue(results and results[-1]["ok"], messages)
            self.assertTrue((folder / "out.xlsx").is_file())
            self.assertTrue((folder / "out.csv").is_file())

    @unittest.skipUnless(os.environ.get("SPECTRA_TEST_GUI") == "1", "set SPECTRA_TEST_GUI=1 for the Tk integration check")
    def test_real_tk_window_stays_responsive_during_xlsx_export(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            folder = Path(raw).resolve()
            files = []
            for file_index in range(3):
                path = folder / f"scan_{file_index}.txt"
                with path.open("w", encoding="utf-8") as handle:
                    for point in range(12000):
                        handle.write(f"{point / 10}\t{point % 197}\n")
                files.append(path)
            output = folder / "gui_export.xlsx"
            try:
                app = sto.SpectraToOriginApp(files)
            except tk.TclError as exc:
                self.skipTest(f"Tk display is unavailable: {exc}")
            try:
                app.root.geometry("860x740-2000-2000")
                app.root.wm_attributes("-alpha", 0.0)
                app.root.deiconify()
                app.root.update()
                app.listbox.selection_set(0)
                app._on_list_selection()
                app.root.update()
                self.assertIn(str(files[0]), app.preview_text.cget("text"))

                height = app.root.winfo_height()
                pending = list(app.root.winfo_children())
                while pending:
                    widget = pending.pop()
                    top = widget.winfo_rooty() - app.root.winfo_rooty()
                    self.assertLessEqual(top + widget.winfo_height(), height + 2, widget)
                    pending.extend(widget.winfo_children())

                ticks = []

                def heartbeat() -> None:
                    ticks.append(time.monotonic())
                    if app._export_process is not None:
                        app.root.after(10, heartbeat)

                with (
                    mock.patch.object(sto.filedialog, "asksaveasfilename", return_value=str(output)),
                    mock.patch.object(sto.messagebox, "showinfo"),
                    mock.patch.object(sto.messagebox, "showerror") as showerror,
                    mock.patch.object(sto.messagebox, "showwarning"),
                ):
                    app.root.after(10, heartbeat)
                    app.export_xlsx()
                    self.assertIsNotNone(app._export_process)
                    deadline = time.monotonic() + 45
                    while app._export_process is not None and time.monotonic() < deadline:
                        app.root.update()
                        time.sleep(0.01)
                    app.root.update()
                    self.assertIsNone(app._export_process, "GUI did not observe worker completion")
                    self.assertGreaterEqual(len(ticks), 2, "Tk after callbacks did not continue during export")
                    showerror.assert_not_called()

                self.assertEqual(app.listbox.cget("state"), "normal")
                self.assertTrue(output.is_file() and output.stat().st_size > 0)
                self.assertTrue(output.with_suffix(".csv").is_file())
                with zipfile.ZipFile(output) as workbook:
                    self.assertIsNone(workbook.testzip())
            finally:
                process = getattr(app, "_export_process", None)
                if process is not None and process.is_alive():
                    process.terminate()
                    process.join(timeout=3)
                app.root.destroy()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    unittest.main()
