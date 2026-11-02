import tempfile
import unittest
from pathlib import Path
from data_cleaner import clean
from xlsx_simple import read_xlsx, write_xlsx


class DataCleanerTests(unittest.TestCase):
    def test_xlsx_round_trip_and_clean(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source, output, report = (root / name for name in ("in.xlsx", "out.xlsx", "report.json"))
            write_xlsx(source, [[" Name ", " Amount "], [" Alice ", " 10 "], ["Alice", "10"], ["", ""], ["Bob", "20"]])
            result = clean(source, output, report)
            self.assertEqual(read_xlsx(output), [["Name", "Amount"], ["Alice", "10"], ["Bob", "20"]])
            self.assertEqual((result["duplicate_rows"], result["blank_rows"]), (1, 1))
            self.assertTrue(report.exists())

    def test_csv_missing_cells_and_malformed(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source, output, report = (root / name for name in ("in.csv", "out.csv", "report.json"))
            source.write_text("a,b\n one \none\nother,two,extra\n", encoding="utf-8")
            result = clean(source, output, report)
            self.assertEqual((result["duplicate_rows"], result["malformed_rows"]), (1, 1))
            self.assertIn("one,", output.read_text(encoding="utf-8-sig"))

    def test_source_not_overwritten(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "in.xlsx"
            write_xlsx(source, [["a"], ["1"]])
            with self.assertRaises(ValueError):
                clean(source, source, Path(folder) / "report.json")

    def test_csv_formula_cells_are_escaped(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source, output, report = (root / name for name in ("in.csv", "out.csv", "report.json"))
            source.write_text("name\n=2+2\n", encoding="utf-8")
            clean(source, output, report)
            self.assertIn("'=2+2", output.read_text(encoding="utf-8-sig"))


if __name__ == "__main__":
    unittest.main()
