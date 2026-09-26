"""Real command-line processes: data discovery, plan replay, and failure protocol."""
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent


class GenericCliTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name).resolve()
        self.source = self.directory / "中文 data.tsv"
        self.source.write_text("Time(s)\tValue(V)\tLabel\n0\t1.2\tA\n1\t\tB\n2\t3.4\tC\n", encoding="utf-8")

    def run_command(self, *args, success=True, launcher=False):
        entry = [str(ROOT / "spectra_to_origin.py"), "agent"] if launcher else ["-m", "origin_bridge"]
        result = subprocess.run([sys.executable, *entry, *map(str, args)], cwd=ROOT,
                                capture_output=True, encoding="utf-8", timeout=60)
        payload = json.loads(result.stdout)
        self.assertEqual(result.returncode, 0 if success else 4, (payload, result.stderr))
        return payload

    def test_inspect_plan_validate_execute_and_refuse_overwrite(self):
        inspected = self.run_command("inspect", "-i", self.source, launcher=True)
        self.assertTrue(inspected["ok"])
        self.assertEqual(len(inspected["tables"]), 1)
        output = self.directory / "result.xlsx"
        saved = self.directory / "import.json"
        plan = self.run_command("plan", "-i", self.source, "-o", output, "--save", saved)
        self.assertEqual(json.loads(saved.read_text(encoding="utf-8")), plan)
        checked = self.run_command("validate", saved)
        self.assertTrue(checked["ok"])
        exported = self.run_command("execute", saved)
        self.assertTrue(exported["ok"])
        book = load_workbook(output)
        try:
            sheet = book.worksheets[0]
            # Header/unit layout is intentionally exporter-owned: verify complete row data.
            values = list(sheet.values)
            self.assertIn((0, 1.2, "A"), values)
            self.assertIn((1, None, "B"), values)
            self.assertIn((2, 3.4, "C"), values)
        finally:
            book.close()
        original = output.read_bytes()
        self.run_command("execute", saved, success=False)
        self.assertEqual(output.read_bytes(), original)

    def test_source_drift_and_invalid_arguments_are_json_failures(self):
        saved = self.directory / "plan.json"
        self.run_command("plan", "-i", self.source, "-o", self.directory / "out.xlsx", "--save", saved)
        self.source.write_text("X,Y\n0,1\n", encoding="utf-8")
        failed = self.run_command("execute", saved, success=False)
        self.assertFalse(failed["ok"])
        self.assertFalse((self.directory / "out.xlsx").exists())
        self.assertFalse(self.run_command("inspect", success=False)["ok"])

    def test_import_table_only_and_relative_plan_paths(self):
        output = self.directory / "out.xlsx"
        plan = self.run_command("plan", "-i", self.source, "-o", output, "--plot", "none")
        plan["tables"][0]["source"]["path"] = self.source.name
        plan["output"]["path"] = output.name
        saved = self.directory / "relative.json"
        saved.write_text(json.dumps(plan), encoding="utf-8")
        self.assertTrue(self.run_command("execute", saved)["ok"])
        self.assertTrue(output.exists())

    def test_launcher_routes_default_and_legacy_gui_inputs(self):
        import spectra_to_origin as sto
        with mock.patch("origin_bridge.gui.GeneralDataApp") as general:
            self.assertEqual(sto.main([str(self.source)]), 0)
            general.assert_called_once_with(initial_files=[self.source])
            general.return_value.run.assert_called_once()
        with mock.patch.object(sto, "SpectraToOriginApp") as legacy:
            self.assertEqual(sto.main(["--spectra-gui"]), 0)
            legacy.return_value.run.assert_called_once()

    def test_validation_exposes_numeric_and_date_conversion_limits_before_write(self):
        self.source.write_text("Time,Value\n2026-01-01T00:00:00.123456789,9007199254740993\n", encoding="utf-8")
        saved = self.directory / "precision.json"
        output = self.directory / "precision.xlsx"
        self.run_command("plan", "-i", self.source, "-o", output, "--save", saved)
        checked = self.run_command("validate", saved)
        self.assertTrue(any("binary64" in warning for warning in checked["warnings"]), checked)
        self.assertTrue(any("milliseconds" in warning for warning in checked["warnings"]), checked)
        self.assertFalse(output.exists())

    def test_value_and_uncertainty_columns_use_row_numbers_as_x(self):
        self.source.write_text("Value(kPa),Error(kPa)\n10,0.2\n12,0.3\n", encoding="utf-8")
        result = self.run_command("inspect", "-i", self.source)
        plot = result["tables"][0]["suggested_plot"]
        self.assertIsNone(plot["x"])
        self.assertEqual(plot["y"], [0])
        self.assertEqual(plot["y_error"], {"Value": 1})


if __name__ == "__main__":
    unittest.main()
