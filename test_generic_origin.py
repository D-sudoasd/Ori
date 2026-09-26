"""Licensed Origin end-to-end import, save, reopen, and data/plot readback."""
import json
import os
import subprocess
import sys
import tempfile
import textwrap
import unittest
from pathlib import Path

import spectra_to_origin as sto


@unittest.skipUnless(sys.platform == "win32" and os.environ.get("SPECTRA_TEST_ORIGIN") == "1",
                     "requires SPECTRA_TEST_ORIGIN=1 and installed Origin")
class GenericOriginRoundTripTests(unittest.TestCase):
    def test_keep_open_uses_final_path_and_existing_session_is_preserved(self):
        if sto.origin_process_running():
            self.skipTest("Origin is already running; preserve the user's session")
        with tempfile.TemporaryDirectory() as raw:
            directory = Path(raw).resolve()
            source = directory / "signal.tsv"
            source.write_text("Time\tSignal\n0\t10\n1\t12\n", encoding="utf-8")
            output = directory / "final.opju"
            command = [sys.executable, "-m", "origin_bridge", "import", "-i", str(source), "-o", str(output)]
            result = subprocess.run([*command, "--keep-open"], capture_output=True, encoding="utf-8", timeout=180)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertTrue(output.exists())
            self.assertTrue(sto.origin_process_running())
            import originpro as op
            try:
                op.attach()
                self.assertEqual(Path(op.path("p")).resolve(), directory)
                self.assertEqual(op.get_lt_str("%G"), output.stem)
                another = directory / "another.opju"
                refused = subprocess.run([*command[:-1], str(another)], capture_output=True,
                                         encoding="utf-8", timeout=60)
                self.assertNotEqual(refused.returncode, 0)
                self.assertFalse(another.exists())
                self.assertEqual(Path(op.path("p")).resolve(), directory)
                self.assertEqual(list(op.pages("w"))[0][0].to_list(1), [10, 12])
            finally:
                sto.close_origin_app(op, started=True)

    def test_mixed_tables_dates_categories_missing_values_and_plot_types(self):
        if sto.origin_process_running():
            self.skipTest("Origin is already running; preserve the user's session")
        from origin_bridge.planning import create_plan, prepare_plan

        with tempfile.TemporaryDirectory() as raw:
            directory = Path(raw).resolve()
            inputs = directory / "inputs"
            inputs.mkdir()
            (inputs / "numeric.tsv").write_text(
                "Time(s)\tForce(N)\tUncertainty(N)\tID\n0\t1.25\t0.1\t001\n1\t\t\t002\n2\t3.75\t0.3\t003\n", encoding="utf-8")
            (inputs / "dates.json").write_text(json.dumps([
                {"Date": "2026-01-01T12:00:00+08:00", "Reading": 11},
                {"Date": "2026-01-02T12:00:00+08:00", "Reading": 12},
                {"Date": "2026-01-03T12:00:00+08:00", "Reading": 13}]), encoding="utf-8")
            (inputs / "category.json").write_text(json.dumps([
                {"Material": "Ti-A", "Strength": 800},
                {"Material": "Ti-B", "Strength": 900},
                {"Material": "Ti-C", "Strength": 850}]), encoding="utf-8")
            (inputs / "row.json").write_text(json.dumps({"Signal": [3, 1, 4]}), encoding="utf-8")
            (inputs / "notes.json").write_text(json.dumps([
                {"Label": "alpha", "Comment": "first"},
                {"Label": "beta", "Comment": None},
                {"Label": "gamma", "Comment": "last"}]), encoding="utf-8")
            output = directory / "general.opju"
            plan = create_plan([inputs], output)
            for table in plan["tables"]:
                name = Path(table["source"]["path"]).stem
                table["name"] = name
                plot = table["plot"]
                plot["title"] = name + " graph"
                if name == "numeric":
                    plot.update(kind="line", x="Time", y=["Force"], y_error={"Force": "Uncertainty"},
                                x_label="Elapsed time (s)", y_label="Force (N)")
                elif name == "dates":
                    plot.update(kind="line_symbol", x="Date", y=["Reading"])
                elif name == "row":
                    plot.update(kind="scatter", x=None, y=["Signal"])
            prepared = prepare_plan(plan)
            saved_plan = directory / "plan.json"
            saved_plan.write_text(json.dumps(plan), encoding="utf-8")
            result = subprocess.run([sys.executable, "-m", "origin_bridge", "execute", str(saved_plan)],
                                    capture_output=True, encoding="utf-8", timeout=240)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            report = json.loads(result.stdout)
            self.assertTrue(report["ok"])
            self.assertTrue(output.exists())
            expected = [{"name": item.name, "columns": [
                {"name": col.name, "kind": col.kind, "unit": col.unit, "values": col.values}
                for col in item.table.columns]} for item in prepared.tables]
            expected_path = directory / "expected.json"
            expected_path.write_text(json.dumps(expected), encoding="utf-8")
            check = textwrap.dedent("""
                import json, math, sys
                from datetime import datetime, timezone
                import spectra_to_origin as sto
                if sto.origin_process_running():
                    raise RuntimeError('Origin is already running')
                import originpro as op
                try:
                    op.set_show(False)
                    assert op.open(sys.argv[1])
                    books = list(op.pages('w'))
                    assert len(books) == 1, len(books)
                    sheets = {sheet.lname: sheet for sheet in books[0]}
                    expected = json.load(open(sys.argv[2], encoding='utf-8'))
                    assert len(sheets) == len(expected) + 1, list(sheets)
                    for table in expected:
                        sheet = sheets[table['name']]
                        assert sheet.cols == len(table['columns']), table['name']
                        for index, col in enumerate(table['columns']):
                            assert sheet.get_label(index, 'L') == col['name']
                            assert sheet.get_label(index, 'U') == col['unit']
                            values = sheet.to_list(index)
                            assert len(values) >= len(col['values']), (table['name'], col, values)
                            for raw, actual in zip(col['values'], values):
                                location = (table['name'], col['name'], raw, actual)
                                if col['kind'] == 'text':
                                    assert actual == (raw or ''), location
                                elif raw is None:
                                    assert math.isnan(actual), location
                                elif col['kind'] == 'number':
                                    assert math.isclose(actual, float(raw), rel_tol=1e-12, abs_tol=1e-12), location
                                else:
                                    dt = datetime.fromisoformat(raw.replace('Z', '+00:00'))
                                    if dt.tzinfo:
                                        dt = dt.astimezone(timezone.utc).replace(tzinfo=None)
                                    days = (dt-datetime(1970,1,1)).total_seconds()/86400
                                    expected_date = days + 25569 + op.lt_float('@DSO')
                                    assert math.isclose(actual, expected_date, rel_tol=0, abs_tol=1e-8), location
                    graphs = list(op.pages('g'))
                    assert len(graphs) == 4, len(graphs)
                    for graph in graphs:
                        assert len(graph[0].plot_list()) >= 1
                    numeric_graph = next(g for g in graphs if 'numeric' in g.lname)
                    assert numeric_graph[0].axis('x').title == 'Elapsed time (s)'
                    assert numeric_graph[0].axis('y').title == 'Force (N)'
                    print('OPJU reopened: all source cells, names, units, dates, categories and four graph types verified')
                finally:
                    sto.close_origin_app(op, started=True)
            """)
            verified = subprocess.run([sys.executable, "-c", check, str(output), str(expected_path)],
                                      capture_output=True, encoding="utf-8", timeout=180)
            self.assertEqual(verified.returncode, 0, verified.stdout + verified.stderr)
            print(verified.stdout.strip())


if __name__ == "__main__":
    unittest.main()
