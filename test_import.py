"""Focused tests for spectrum text parsing and file collection."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import spectra_to_origin as sto


class SpectrumParsingTests(unittest.TestCase):
    def test_whitespace_tab_comma_and_semicolon_with_optional_headers(self) -> None:
        rows = {
            "space.txt": "2theta intensity\n0.10 12.5\n0.20 13.5\n",
            "tab.tsv": "angle\tcounts\n0.10\t12.5\n0.20\t13.5\n",
            "comma.csv": "x,y\n0.10,12.5\n0.20,13.5\n",
            "semicolon.dat": "x;signal\n0.10;12.5\n0.20;13.5\n",
        }
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            for name, text in rows.items():
                path = root / name
                path.write_text(text, encoding="utf-8")
                with self.subTest(name=name):
                    spectrum = sto.parse_spectrum(path)
                    self.assertEqual(spectrum.x_text, ("0.10", "0.20"))
                    self.assertEqual(spectrum.y_text, ("12.5", "13.5"))

    def test_supported_encodings_and_d_exponents_preserve_decimal_text(self) -> None:
        samples = [
            ("utf8.txt", "utf-8-sig", "# eta=0.5\nX,Y\n+001.2300D-04,5d+2\n2.00,6\n", ("+001.2300E-04", "2.00"), ("5e+2", "6")),
            ("gbk.dat", "gbk", "角度 强度\n0.0100 4.00\n0.0200 5.00\n", ("0.0100", "0.0200"), ("4.00", "5.00")),
            ("utf16.xy", "utf-16", "x\ty\n1.2500\t2.5000\n2.2500\t3.5000\n", ("1.2500", "2.2500"), ("2.5000", "3.5000")),
        ]
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            for name, encoding, text, x_expected, y_expected in samples:
                path = root / name
                path.write_text(text, encoding=encoding)
                with self.subTest(name=name):
                    spectrum = sto.parse_spectrum(path)
                    self.assertEqual(spectrum.x_text, x_expected)
                    self.assertEqual(spectrum.y_text, y_expected)

    def test_utf16_without_bom_is_not_accepted_as_bom_marked_utf16(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            path = Path(raw) / "utf16_no_bom.txt"
            path.write_bytes("0\t1\n1\t2\n".encode("utf-16-le"))
            with self.assertRaises(ValueError):
                sto.parse_spectrum(path)

    def test_nonfinite_values_report_filename_and_physical_line(self) -> None:
        values = ["NaN", "-Infinity", "1e999"]
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            for value in values:
                path = root / "bad.txt"
                path.write_text(f"# comment\n0 1\n1 {value}\n2 3\n", encoding="utf-8")
                with self.subTest(value=value):
                    with self.assertRaisesRegex(ValueError, r"bad\.txt 第 3 行包含非有限数值"):
                        sto.parse_spectrum(path)

    def test_bad_numeric_row_reports_physical_line(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            path = Path(raw) / "bad.csv"
            path.write_text("x,y\n0,1\n1,noise\n", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, r"bad\.csv 第 3 行不是数值"):
                sto.parse_spectrum(path)

    def test_first_bad_data_row_is_rejected_instead_of_treated_as_a_header(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            path = Path(raw) / "bad_first.csv"
            path.write_text("0,noise\n1,2\n2,3\n", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, r"bad_first\.csv 第 1 行不是数值"):
                sto.parse_spectrum(path)

    def test_text_row_after_numeric_data_is_not_silently_skipped_as_a_header(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            path = Path(raw) / "bad_middle.csv"
            path.write_text("0,1\n1,2\nnoise,error\n2,3\n", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, r"bad_middle\.csv 第 3 行不是数值"):
                sto.parse_spectrum(path)

    def test_header_values_may_contain_digits_and_csv_fields_may_be_quoted(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            path = Path(raw) / "quoted.csv"
            path.write_text(
                '"2theta, deg","I4000 (a.u.)"\n"0.10","12.5"\n"0.20","13.5"\n',
                encoding="utf-8",
            )
            spectrum = sto.parse_spectrum(path)
            self.assertEqual(spectrum.x_text, ("0.10", "0.20"))
            self.assertEqual(spectrum.y_text, ("12.5", "13.5"))

    def test_csv_quote_error_reports_filename_and_physical_line(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            path = Path(raw) / "bad_quote.csv"
            path.write_text('0,1\n1,"2\n2,3\n', encoding="utf-8")
            with self.assertRaisesRegex(ValueError, r"bad_quote\.csv 第 2 行分隔符格式无效"):
                sto.parse_spectrum(path)

    def test_nonfinite_first_row_is_not_mistaken_for_a_header(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            path = Path(raw) / "bad.tsv"
            path.write_text("NaN\tInf\n1\t2\n", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, r"bad\.tsv 第 1 行包含非有限数值"):
                sto.parse_spectrum(path)


class CollectionTests(unittest.TestCase):
    def test_supported_suffixes_are_collected_in_natural_order(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            for name in ("scan10.txt", "scan2.csv", "scan1.xy", "ignore.pdf", "scan3.tsv", "scan4.dat"):
                (root / name).write_text("0 1\n1 2\n", encoding="utf-8")
            result = sto.collect_txt_files(root)
            self.assertEqual(
                [path.name for path in result],
                ["scan1.xy", "scan2.csv", "scan3.tsv", "scan4.dat", "scan10.txt"],
            )

    def test_recursive_collection_sorts_naturally_and_includes_temp_copy_for_comparison(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            (root / "scan10.txt").write_text("0 10\n1 11\n", encoding="utf-8")
            (root / "scan2.txt").write_text("0 2\n1 3\n", encoding="utf-8")
            temp_copy = root / sto.TEMP_COPY_DIR
            temp_copy.mkdir()
            (temp_copy / "scan10.txt").write_text("0 10\n1 11\n", encoding="utf-8")
            result = sto.collect_txt_files(root)
            self.assertEqual(
                [path.relative_to(root).as_posix() for path in result],
                ["scan2.txt", "scan10.txt", f"{sto.TEMP_COPY_DIR}/scan10.txt"],
            )

    def test_merge_keeps_distinct_same_name_files_and_only_removes_real_temp_copy(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            first = root / "series-a" / "scan.txt"
            second = root / "series-b" / "scan.txt"
            original = root / "scan.txt"
            copy = root / sto.TEMP_COPY_DIR / "scan.txt"
            for path in (first, second, original, copy):
                path.parent.mkdir(parents=True, exist_ok=True)
            first.write_text("0 1\n1 2\n", encoding="utf-8")
            second.write_text("0 9\n1 8\n", encoding="utf-8")
            original.write_text("0 1\n1 2\n", encoding="utf-8")
            copy.write_text("0 1\n1 2\n", encoding="utf-8")

            merged = sto.merge_file_list([], [first, second])
            self.assertEqual(merged, [first, second])
            merged = sto.merge_file_list(merged, [copy, original])
            self.assertEqual(merged, [first, second, original])
            self.assertNotIn(copy, merged)

    def test_same_name_regular_files_are_retained_even_when_bytes_match(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            first = root / "one" / "same.txt"
            second = root / "two" / "same.txt"
            first.parent.mkdir()
            second.parent.mkdir()
            first.write_text("0 1\n1 2\n", encoding="utf-8")
            second.write_bytes(first.read_bytes())
            self.assertEqual(sto.merge_file_list([], [first, second]), [first, second])

    def test_same_bytes_in_unrelated_trees_are_not_treated_as_temp_copies(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            temp_named = root / "experiment-a" / sto.TEMP_COPY_DIR / "scan.txt"
            independent = root / "experiment-b" / "scan.txt"
            temp_named.parent.mkdir(parents=True)
            independent.parent.mkdir(parents=True)
            temp_named.write_text("0 1\n1 2\n", encoding="utf-8")
            independent.write_bytes(temp_named.read_bytes())
            self.assertEqual(
                sto.merge_file_list([], [temp_named, independent]),
                [temp_named, independent],
            )

    def test_exact_sibling_temp_copy_is_deduplicated_and_prefers_primary(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            primary = root / "scan.txt"
            copy = root / sto.TEMP_COPY_DIR / "scan.txt"
            copy.parent.mkdir()
            primary.write_text("0 1\n1 2\n", encoding="utf-8")
            copy.write_bytes(primary.read_bytes())
            self.assertEqual(sto.merge_file_list([], [copy, primary]), [primary])

    def test_user_path_collection_returns_invalid_inputs_as_errors(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            valid = root / "valid.csv"
            invalid = root / "missing.txt"
            valid.write_text("0,1\n1,2\n", encoding="utf-8")
            files, errors = sto.collect_from_user_paths([valid, invalid])
            self.assertEqual(files, [valid.resolve()])
            self.assertEqual(len(errors), 1)
            self.assertIn(str(invalid), errors[0])


if __name__ == "__main__":
    unittest.main()
