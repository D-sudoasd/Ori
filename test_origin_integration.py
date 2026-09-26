"""Opt-in tests against a licensed, installed Origin, with numeric readback."""

import os
import subprocess
import sys
import tempfile
import textwrap
import unittest
from pathlib import Path

import spectra_to_origin as sto


@unittest.skipUnless(sys.platform == "win32" and os.environ.get("SPECTRA_TEST_ORIGIN") == "1", "requires SPECTRA_TEST_ORIGIN=1 and installed Origin")
class OriginGroupRoundTripTests(unittest.TestCase):
    def test_mixed_layouts_custom_axes_and_unequal_lengths(self):
        if sto.origin_process_running():
            self.skipTest("Origin is already running; preserve the user's session")
        with tempfile.TemporaryDirectory() as raw:
            folder = Path(raw)
            source = folder / "input"
            source.mkdir()
            (source / "400C_sample1.csv").write_text("x,y\n0,10\n1,11\n2,12\n", encoding="utf-8")
            (source / "400C_sample2.txt").write_text("0.0 20\n1.0 21\n2.0 22\n", encoding="utf-8")
            (source / "500C_sample1.dat").write_text("x;y\n10;30\n11;31\n12;32\n", encoding="utf-8")
            (source / "500C_sample2.xy").write_text("15 4D+1\n16 4.1D+1\n", encoding="utf-8")
            target = folder / "mixed.opju"
            command = [sys.executable, str(Path(sto.__file__).resolve()), "--cli", "-i", str(source), "-o", str(target), "--group-by-name", "--also-xlsx", "--x-name", "q", "--x-unit", "nm^-1", "--y-name", "Counts", "--y-unit", "counts"]
            exported = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180)
            self.assertEqual(exported.returncode, 0, exported.stdout + exported.stderr)
            self.assertTrue(target.is_file())
            script = textwrap.dedent("""
                import math
                import sys
                import spectra_to_origin as sto
                if sto.origin_process_running():
                    raise RuntimeError('An Origin session is already open')
                import originpro as op
                try:
                    op.set_show(False)
                    assert op.open(sys.argv[1])
                    book = list(op.pages('w'))[0]
                    assert book[0].cols == 3
                    assert book[1].cols == 4
                    assert book[0].to_list(0) == [0, 1, 2]
                    assert book[0].to_list(1) == [10, 11, 12]
                    assert book[0].to_list(2) == [20, 21, 22]
                    assert book[1].to_list(0) == [10, 11, 12]
                    assert book[1].to_list(1) == [30, 31, 32]
                    assert book[1].to_list(2)[:2] == [15, 16]
                    assert book[1].to_list(3)[:2] == [40, 41]
                    for value in book[1].to_list(3)[2:]:
                        assert math.isnan(value)
                    assert book[0].get_label(0, 'L') == 'q'
                    assert book[0].get_label(0, 'U') == 'nm^-1'
                    assert book[1].get_label(3, 'U') == 'counts'
                    graphs = list(op.pages('g'))
                    assert len(graphs) == 2
                    for graph in graphs:
                        assert len(list(graph[0].obj.DataPlots)) == 2
                        assert graph[0].axis('x').title == 'q (nm^-1)'
                        assert graph[0].axis('y').title == 'Counts (counts)'
                    print('MIXED_OPJU numeric values, labels, axes and plots verified')
                finally:
                    sto.close_origin_app(op, started=True)
            """)
            checked = subprocess.run([sys.executable, "-c", script, str(target)], cwd=Path(sto.__file__).parent, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=120)
            self.assertEqual(checked.returncode, 0, checked.stdout + checked.stderr)
            self.assertTrue(target.with_suffix(".xlsx").is_file())
            self.assertEqual(len(list((folder / "mixed_csv").glob("*.csv"))), 2)
            print(checked.stdout.strip())


if __name__ == "__main__":
    unittest.main()
