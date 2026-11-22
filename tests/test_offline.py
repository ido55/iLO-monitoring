"""Offline tests for the iLO monitor; none of these tests contact a server."""
import csv
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd

PROJECT = Path(__file__).resolve().parents[1]
SAMPLE = PROJECT / "ScriptFiles" / "sample-response.xml"
sys.path.insert(0, str(PROJECT))
import GUI_iLO_Script as app


class OfflineChecks(unittest.TestCase):
    def test_sample_response_parses_health_and_identity(self):
        row = app.parse_health_response(SAMPLE.read_text(encoding="utf-8-sig"), "sample-target", "Sample site")
        self.assertEqual(row["Server Name"], "sample-server")
        self.assertEqual(row["iLO Version"], "2.99")
        self.assertEqual(row["Fans Status"], "OK")
        self.assertEqual(row["Fans Redundancy"], "Redundant")
        self.assertEqual(row["Collection Error"], "")

    def test_malformed_response_returns_clear_error(self):
        row = app.parse_health_response("<RIBCL>", "sample-target", "Sample site")
        self.assertTrue(row["Collection Error"].startswith("Invalid XML:"))

    def test_csv_export_fallback_writes_sample_row(self):
        row = app.parse_health_response(SAMPLE.read_text(encoding="utf-8-sig"), "sample-target", "Sample site")
        with tempfile.TemporaryDirectory() as temp_dir:
            app.OUTPUT_DIR = Path(temp_dir)
            app.OUTPUT_FILE = Path(temp_dir) / "Output.xlsx"
            app.results = [row]
            with patch.object(pd, "ExcelWriter", side_effect=ImportError("test CSV fallback")):
                with patch.object(app.messagebox, "showinfo"):
                    app.Text_To_Excel()
            csv_path = Path(temp_dir) / "Output.csv"
            with csv_path.open(newline="", encoding="utf-8-sig") as handle:
                saved = next(csv.DictReader(handle))
            self.assertEqual(saved["IP address"], "sample-target")
            self.assertEqual(saved["Fans Redundancy"], "Redundant")


if __name__ == "__main__":
    unittest.main()
