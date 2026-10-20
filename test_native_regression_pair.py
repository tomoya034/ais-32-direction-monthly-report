from __future__ import annotations

import copy
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from scripts.compare_native_regression_pair import GEOMETRY_FIELDS, _native_value, compare, sha256, validated_capture


class NativeRepeatContracts(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.root = Path(self.directory.name)
        self.report = {"state": "COMPLETE", "root": str(self.root),
                       "expected_workbook_count": 6, "workbook_count": 6,
                       "pass_count": 6, "fail_count": 0, "workbooks": []}
        for number in range(6):
            workbook = self.root / f"book_{number}.xlsx"
            workbook.write_bytes(b"synthetic contract fixture, not an Excel file")
            charts = []
            for index in (1, 2):
                png = self.root / f"image_{number}_{index}.png"
                png.write_bytes(b"synthetic inventory placeholder; pixel tests use actual images")
                chart = {field: 1.0 for field in GEOMETRY_FIELDS}
                chart.update({"title": "fixture", "sheet": "Summary", "index": index,
                              "png": str(png), "png_exported": True,
                              "png_validation": {"decoded": True, "status": "VALID_PNG",
                                                 "sha256": sha256(png)}})
                charts.append(chart)
            self.report["workbooks"].append({"path": str(workbook), "status": "PASS",
                    "opened_read_only": True, "file_unchanged": True, "error": None,
                    "sha256_before": sha256(workbook), "sha256_after": sha256(workbook),
                    "formula_errors": [], "sheets": ["Summary"],
                    "formula_checks": [{"sheet": "Summary", "result": {
                        "status": "VERIFIED_ZERO_ERRORS", "error_count": 0,
                        "formula_cell_count": 10, "error": None}}], "charts": charts})

    def tearDown(self):
        self.directory.cleanup()

    def add_missing_gate(self):
        book = self.report["workbooks"][0]
        book["sheets"] = ["總表"]
        labels = [f"fixture_label_{index}" for index in range(32)]
        values = [0.0] + [None]*31
        helpers = [{"cell": f"{column}5", "source": f"B{row}",
                    "formula": f"=IF(COUNT(B{row})=0,NA(),B{row})",
                    "expected_value": 0.0, "expected_missing": False}
                   for column, row in zip(("AJ", "AK", "AL"), (7, 8, 9))]
        radar = []
        for index in range(32):
            # AO=41 through BT=72; contract fixtures do not generate a workbook.
            def column(number):
                result = ""
                while number:
                    number, offset = divmod(number-1, 26)
                    result = chr(65+offset) + result
                return result
            source = f"{column(index+2)}5"
            radar.append({"cell": f"{column(index+41)}5", "source": source,
                          "formula": f"=IF(COUNT({source})=0,NA(),{source})",
                          "expected_value": values[index], "expected_missing": index != 0})
        cache = self.root / "native_contract_source.cache"
        cache.write_bytes(b"test fixture: independent source binding")
        contract = {"schema": 1, "role": "Modern/default", "state": "PASS_SERIALIZATION_PREREQUISITES",
                    "workbook": book["path"], "workbook_sha256": book["sha256_before"],
                    "cache_sources": [{"path": str(cache), "sha256": sha256(cache)}],
                    "helpers": helpers, "radar_helpers": radar, "expected_matrix": [values],
                    "included_indices": [0], "direction_labels": labels, "logical_days": ["2026-01-01"]}
        contract_path = self.root / "native_contract.json"
        contract_path.write_text(json.dumps({"modern_chart_missing_contract": contract}), encoding="utf-8")
        observations = [{"cell": item["cell"], "source": item["source"], "formula": item["formula"],
                         "expected_missing": item["expected_missing"],
                         "source_count": 0 if item["expected_missing"] else 1,
                         "source_value": "" if item["expected_missing"] else 0.0,
                         "helper_value": -2146826246 if item["expected_missing"] else 0.0}
                        for item in helpers+radar]
        allowed = [item["cell"] for item in radar[1:]]
        series = [{"chart_index": 1, "series_index": 1, "name": "1/1",
                   "formula": '=SERIES("1/1",總表!$B$4:$AG$4,總表!$AO$5:$BT$5,1)',
                   "point_count": 32, "categories": labels, "values": values.copy(),
                   "missing_native_null_points": 31, "genuine_zero_points": 1}]
        for index, (name, helper) in enumerate(zip(("最遠", "最近", "平均"), helpers), 1):
            column = helper["cell"][:2]
            series.append({"chart_index": 2, "series_index": index, "name": name,
                           "formula": f'=SERIES("{name}",總表!$AI$5:$AI$5,總表!${column}$5:${column}$5,{index})',
                           "point_count": 1, "categories": [labels[0]], "values": [0.0],
                           "missing_native_null_points": 0, "genuine_zero_points": 1})
        book["modern_chart_missing"] = {"status": "PASS", "independent_contract_status": "PASS", "error": None,
            "contract_report": str(contract_path), "contract_report_sha256": sha256(contract_path),
            "workbook_sha256": book["sha256_before"], "intentional_error_addresses": allowed,
            "helper_observations": observations, "primary_daily_cells": 32,
            "statistic_helper_cells": 3, "radar_helper_cells": 32, "series_observations": series,
            "series_points": 35, "genuine_zero_points": 4}
        book["formula_checks"] = [{"sheet": "總表", "result": {
            "status": "VERIFIED_INTENTIONAL_CHART_MISSING", "error_count": 31,
            "formula_cell_count": 100, "error": None, "unexpected_error_count": 0,
            "intentional_chart_missing_count": 31, "intentional_chart_missing_addresses": allowed}}]
        return book

    def test_precise_native_intentional_markers_accepted(self):
        self.add_missing_gate()
        self.assertEqual(len(validated_capture(self.report)), 6)

    def test_native_missing_zero_wrong_scope_and_false_healthy_zero_rejected(self):
        self.add_missing_gate()
        for case in ("missing_zero", "genuine_zero_null", "wrong_scope", "false_zero", "unexpected_error", "missing_gate", "shifted_series", "incomplete_helpers"):
            with self.subTest(case=case):
                report = copy.deepcopy(self.report)
                book = report["workbooks"][0]
                gate = book["modern_chart_missing"]
                if case == "missing_zero":
                    gate["series_observations"][0]["values"][31] = 0.0
                elif case == "genuine_zero_null":
                    gate["series_observations"][0]["values"][0] = None
                elif case == "wrong_scope":
                    gate["intentional_error_addresses"][0] = "B7"
                elif case == "false_zero":
                    book["formula_checks"][0]["result"]["error_count"] = 0
                elif case == "unexpected_error":
                    book["formula_checks"][0]["result"]["unexpected_error_count"] = 1
                elif case == "missing_gate":
                    book["modern_chart_missing"] = None
                elif case == "shifted_series":
                    gate["series_observations"][0]["formula"] = '=SERIES("1/1",總表!$B$4:$AG$4,總表!$B$5:$AG$5,1)'
                else:
                    gate["helper_observations"].pop()
                with self.assertRaises(ValueError):
                    validated_capture(report)

    def test_native_contract_and_cache_must_remain_current(self):
        book = self.add_missing_gate()
        contract = Path(book["modern_chart_missing"]["contract_report"])
        previous = contract.read_bytes()
        contract.write_bytes(previous + b" ")
        with self.assertRaises(ValueError):
            validated_capture(self.report)
        contract.write_bytes(previous)
        (self.root / "native_contract_source.cache").write_bytes(b"changed")
        with self.assertRaises(ValueError):
            validated_capture(self.report)

    def test_native_zero_and_missing_are_exact_not_tolerance_substitutes(self):
        _native_value(0.0, 0.0)
        _native_value(None, None)
        for actual, expected in ((5e-13, 0.0), (0.0, None), (None, 0.0), (True, 0.0), (-2146826246, 0.0), ("#N/A", None)):
            with self.assertRaises(ValueError):
                _native_value(actual, expected)

    def test_all_six_current_workbooks_are_covered(self):
        self.assertEqual(len(validated_capture(self.report)), 6)

    def test_missing_and_extra_workbooks_rejected(self):
        missing = copy.deepcopy(self.report)
        missing["workbooks"].pop()
        with self.assertRaises(ValueError):
            validated_capture(missing)
        (self.root / "unexpected.xlsx").write_bytes(b"unexpected fixture")
        with self.assertRaises(ValueError):
            validated_capture(self.report)

    def test_zero_one_and_duplicate_charts_rejected(self):
        for length in (0, 1, 3):
            altered = copy.deepcopy(self.report)
            altered["workbooks"][0]["charts"] = (altered["workbooks"][0]["charts"] * 2)[:length]
            with self.assertRaises(ValueError):
                validated_capture(altered)

    def test_unknown_formula_scan_and_omitted_sheet_rejected(self):
        for edit in ("unknown", "omitted", "false_zero"):
            altered = copy.deepcopy(self.report)
            if edit == "unknown":
                altered["workbooks"][0]["formula_checks"][0]["result"]["status"] = "UNVERIFIED"
            elif edit == "omitted":
                altered["workbooks"][0]["formula_checks"] = []
            else:
                altered["workbooks"][0]["formula_checks"][0]["result"]["formula_cell_count"] = 0
            with self.assertRaises(ValueError):
                validated_capture(altered)

    def test_changed_current_workbook_and_png_rejected(self):
        book = Path(self.report["workbooks"][0]["path"])
        book.write_bytes(b"changed")
        with self.assertRaises(ValueError):
            validated_capture(self.report)
        book.write_bytes(b"synthetic contract fixture, not an Excel file")
        Path(self.report["workbooks"][0]["charts"][0]["png"]).write_bytes(b"changed PNG")
        with self.assertRaises(ValueError):
            validated_capture(self.report)

    def test_identical_pairs_and_shifted_native_geometry(self):
        with patch("scripts.compare_native_regression_pair.compare_png", return_value={"status": "EXACT_EQUAL", "changed_pixels": 0}):
            result = compare(self.report, self.report)
            self.assertEqual((result["status"], result["chart_pairs"], result["pass_count"]), ("PASS", 12, 12))
            changed = copy.deepcopy(self.report)
            changed["workbooks"][0]["charts"][0]["plot_width_points"] += 1
            self.assertEqual(compare(self.report, changed)["status"], "FAIL")

    def test_pixels_are_required_even_when_geometry_matches(self):
        with patch("scripts.compare_native_regression_pair.compare_png", return_value={"status": "PIXEL_DIFFERENCE", "changed_pixels": 1}):
            self.assertEqual(compare(self.report, self.report)["status"], "FAIL")


if __name__ == "__main__":
    unittest.main()
