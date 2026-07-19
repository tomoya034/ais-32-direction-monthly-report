from __future__ import annotations

import tempfile
import unittest
import shutil
from datetime import date
from pathlib import Path

import openpyxl

from ais_monthly_app import (
    AppConfig,
    Candidate,
    derive_legacy_output_path,
    degree_to_direction_index,
    friendly_error_message,
    parse_source_date,
    run_pipeline,
    select_cluster,
    select_cluster_from_sorted_rows,
)


class RuleTests(unittest.TestCase):
    def test_filename_date_uses_yyyymmdd_component(self) -> None:
        self.assertEqual(parse_source_date("D&TMOK KLNG_20260401_23.xlsx"), date(2026, 4, 1))
        self.assertIsNone(parse_source_date("KLNG_20260401.xlsx"))

    def test_direction_mapping_matches_existing_script_bins(self) -> None:
        self.assertEqual(degree_to_direction_index(0)[0], 0)
        self.assertEqual(degree_to_direction_index(11.25)[0], 1)
        self.assertEqual(degree_to_direction_index(359.9)[0], 31)
        self.assertEqual(degree_to_direction_index(360)[0], 0)

    def test_cluster_skips_isolated_high_value(self) -> None:
        values = [100.0, 80.0, 79.0, 78.0, 60.0]
        candidates = [Candidate(value, 0.0, index + 2) for index, value in enumerate(values)]
        result = select_cluster("北", candidates, tolerance=0.10, cluster_size=3, over_cap_count=0, over_cap_max=None)
        self.assertEqual(result.selected, 80.0)
        self.assertEqual(result.selected_rank, 2)
        self.assertEqual(result.status, "自動採用")

    def test_coastal_value_is_flagged(self) -> None:
        candidates = [Candidate(value, 247.5, index + 2) for index, value in enumerate((12.0, 11.8, 11.5))]
        result = select_cluster("西南西", candidates, tolerance=0.10, cluster_size=3, over_cap_count=0, over_cap_max=None)
        self.assertEqual(result.status, "待複核")

    def test_full_numeric_search_can_find_cluster_below_saved_top_rows(self) -> None:
        config = AppConfig(
            input_dir=Path("."),
            output_path=Path("result.xlsx"),
            year=2026,
            month=4,
            top_candidates=3,
        )
        rows = [(100.0, 2, 0.0), (80.0, 3, 0.0), (60.0, 4, 0.0), (40.0, 5, 0.0), (20.0, 6, 0.0), (19.0, 7, 0.0), (18.0, 8, 0.0)]
        result = select_cluster_from_sorted_rows("北", rows, config, over_cap_count=0, over_cap_max=None)
        self.assertEqual(result.selected, 20.0)
        self.assertEqual(result.selected_rank, 5)
        self.assertEqual([candidate.rank for candidate in result.candidates], [1, 2, 3, 5, 6, 7])

    def test_output_names_and_common_error_are_friendly(self) -> None:
        modern = Path("KLNG 2026年4月 32方位數值_新版自動分析.xlsx")
        self.assertEqual(derive_legacy_output_path(modern).name, "KLNG 2026年4月 32方位數值_原格式相容版.xlsx")
        self.assertIn("Excel", friendly_error_message(PermissionError("locked")))


class EndToEndTests(unittest.TestCase):
    def test_small_month_build(self) -> None:
        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as temporary:
            folder = Path(temporary)
            source = folder / "D&TMOK KLNG_20260401_23.xlsx"
            workbook = openpyxl.Workbook(write_only=True)
            sheet = workbook.create_sheet("AIS")
            sheet.append(["msg_type", "LONGITUDE_DESC", "bearing", "distance in nautical miles", "unused"])
            rows = [
                (1, "East", 0.0, 100.0),
                (3, "East", 0.0, 80.0),
                (18, "East", 0.0, 79.0),
                (19, "East", 0.0, 78.0),
                (2, "East", 247.5, 12.0),
                (2, "East", 247.5, 11.8),
                (2, "East", 247.5, 11.5),
                (2, "East", 0.0, 499.0),
                (1, "West", 0.0, 499.0),
                (1, "East", 0.0, 600.0),
            ]
            for row in rows:
                sheet.append([*row, "x"])
            workbook.save(source)
            shutil.copyfile(source, folder / "D&TMOK KLNG_20260402_23.xlsx")

            output = folder / "result.xlsx"
            legacy_output = folder / "result_legacy.xlsx"
            config = AppConfig(
                input_dir=folder,
                output_path=output,
                year=2026,
                month=4,
                legacy_output_path=legacy_output,
                top_candidates=10,
                workers=2,
                overwrite=True,
            )
            run_pipeline(config)
            self.assertTrue(output.exists())
            self.assertTrue(legacy_output.exists())

            events = []
            run_pipeline(config, callback=events.append)
            self.assertTrue(any(event.get("kind") == "cache_hit" for event in events))

            result = openpyxl.load_workbook(output, read_only=False, data_only=False)
            try:
                self.assertIn("總表", result.sheetnames)
                self.assertIn("待複核", result.sheetnames)
                self.assertEqual(result["4月1日"]["H2"].value, 80.0)
                self.assertEqual(result["4月1日"]["AD2"].value, 12.0)
                self.assertIn("'4月1日'!H4", str(result["總表"]["B5"].value))
                self.assertEqual(len(result["總表"]._charts), 2)
            finally:
                result.close()

            legacy = openpyxl.load_workbook(legacy_output, read_only=False, data_only=False)
            try:
                self.assertIn("4月1日", legacy.sheetnames)
                self.assertIn("總表", legacy.sheetnames)
                self.assertIn("工作", legacy.sheetnames)
                self.assertEqual(legacy["4月1日"]["A2"].value, 600.0)
                self.assertEqual(legacy["4月1日"]["C2"].value, "北")
                self.assertEqual(legacy["4月1日"]["H2"].value, 80.0)
                self.assertEqual(legacy["4月1日"]["AD2"].value, 12.0)
                self.assertIn("'4月1日'!H2", str(legacy["總表"]["B2"].value))
                self.assertIsNone(legacy["總表"]["M2"].value)
                self.assertIsNone(legacy["總表"]["M33"].value)
                self.assertEqual(legacy["工作"].sheet_state, "hidden")
            finally:
                legacy.close()


if __name__ == "__main__":
    unittest.main()
