from __future__ import annotations

import datetime as dt
import json
import math
import os
import unittest
from pathlib import Path

import openpyxl

from ais_monthly_app import (
    AppConfig,
    DayInput,
    DIRECTION_ORDER,
    LEGACY_DIRECTION_ABBREVIATIONS,
    OPEN_SEA_INDEXES,
    PERIOD_A,
    PERIOD_B,
    SourceFileError,
    SourceFragment,
    _display_candidate_catalog,
    _normalized_spool_path,
    build_default_decision_snapshot,
    build_processing_job,
    parse_source_filename,
    process_day_input,
    read_normalized_spool,
    read_review_decisions,
    resolve_decision_values,
    scan_source_files,
    select_cluster_from_sorted_rows,
    source_manifest_hash,
    write_monthly_workbook,
)


def _read_legacy_grid(path: Path) -> dict[tuple[dt.date, int], float | None]:
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook.worksheets[0]
        rows = sheet.iter_rows(values_only=True)
        headers = next(rows)
        abbreviation_to_index = {
            abbreviation: index
            for index, abbreviation in enumerate(LEGACY_DIRECTION_ABBREVIATIONS)
        }
        grid: dict[tuple[dt.date, int], float | None] = {}
        for row in rows:
            raw_date = row[0] if row else None
            if isinstance(raw_date, dt.datetime):
                day = raw_date.date()
            elif isinstance(raw_date, dt.date):
                day = raw_date
            else:
                continue
            for column, header in enumerate(headers[1:], start=1):
                direction_index = abbreviation_to_index.get(str(header))
                if direction_index is None:
                    continue
                value = row[column] if column < len(row) else None
                grid[(day, direction_index)] = None if value is None else float(value)
        return grid
    finally:
        workbook.close()


class SelectorBaselineTests(unittest.TestCase):
    def test_full_search_result_does_not_depend_on_display_limit(self) -> None:
        rows = [
            (100.0, 2, 0.0),
            (80.0, 3, 0.0),
            (60.0, 4, 0.0),
            (40.0, 5, 0.0),
            (20.0, 6, 0.0),
            (19.0, 7, 0.0),
            (18.0, 8, 0.0),
        ]
        selected = []
        for display_limit in (3, 5, 50):
            config = AppConfig(
                input_dir=Path("."),
                output_path=Path("result.xlsx"),
                port="KLNG",
                year=2026,
                month=6,
                top_candidates=display_limit,
            )
            result = select_cluster_from_sorted_rows(
                DIRECTION_ORDER[0], rows, config, over_cap_count=0, over_cap_max=None
            )
            selected.append((result.selected, result.selected_rank, result.cluster_count))
        self.assertEqual(selected, [(20.0, 5, 3)] * 3)


class FragmentCatalogTests(unittest.TestCase):
    @staticmethod
    def _write_fragment(path: Path, distances: list[float]) -> None:
        parsed = parse_source_filename(path.name)
        if parsed is None:
            raise AssertionError(path)
        hour = 13 if parsed.suffix == "23" else 1
        workbook = openpyxl.Workbook(write_only=True)
        sheet = workbook.create_sheet("AIS")
        sheet.append(["Year", "Month", "Day", "Hour", "Minute", "Second", "channel", "msg_type", "mmsi", "LONGITUDE_DESC", "bearing", "distance in nautical miles"])
        for second, distance in enumerate(distances):
            sheet.append([parsed.day.year, parsed.day.month, parsed.day.day, hour, 0, second, "A", 1, 400000000 + second, "East", 0.0, distance])
        workbook.save(path)

    @staticmethod
    def _write_records(path: Path, records: list[tuple[int, int, str, int, float]]) -> None:
        parsed = parse_source_filename(path.name)
        if parsed is None:
            raise AssertionError(path)
        workbook = openpyxl.Workbook(write_only=True)
        sheet = workbook.create_sheet("AIS")
        sheet.append(["Year", "Month", "Day", "Hour", "Minute", "Second", "channel", "msg_type", "mmsi", "LONGITUDE_DESC", "bearing", "distance in nautical miles"])
        for hour, second, channel, msg_type, distance in records:
            sheet.append([parsed.day.year, parsed.day.month, parsed.day.day, hour, 0, second, channel, msg_type, 500000000 + second, "East", 0.0, distance])
        workbook.save(path)

    def test_suffix_is_opaque_traceability_not_source_identity(self) -> None:
        first = parse_source_filename("D&TMOK KLNG_20260601_11.xlsx")
        second = parse_source_filename("D&TMOK KLNG_20260601_part-anything.xlsx")
        self.assertIsNotNone(first)
        self.assertIsNotNone(second)
        self.assertEqual((first.port, first.day), (second.port, second.day))
        self.assertEqual(first.suffix, "11")
        self.assertEqual(second.suffix, "PART-ANYTHING")

    def test_same_day_fragments_are_all_kept_and_union_selected(self) -> None:
        import tempfile

        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as temporary:
            folder = Path(temporary)
            first_path = folder / "D&TMOK KLNG_20260601_11.xlsx"
            second_path = folder / "D&TMOK KLNG_20260601_23.xlsx"
            self._write_fragment(first_path, [100.0, 99.0])
            self._write_fragment(second_path, [98.0])

            catalog, warnings = scan_source_files(folder)
            day_input = catalog["KLNG"][dt.date(2026, 6, 1)]
            self.assertEqual(
                [fragment.path.name for fragment in day_input.fragments],
                [first_path.name, second_path.name],
            )
            self.assertTrue(any("合併 2 個來源分片" in warning for warning in warnings["KLNG"]))

            config = AppConfig(
                input_dir=folder,
                output_path=folder / "result.xlsx",
                port="KLNG",
                year=2026,
                month=6,
                top_candidates=3,
            )
            result = process_day_input(day_input, config)
            self.assertEqual(result.source_fragments, (first_path, second_path))
            self.assertIsNone(result.source_file)
            self.assertEqual(result.rows_scanned, 3)
            self.assertEqual(result.directions["北"].selected, 100.0)

            job = build_processing_job(config)
            self.assertEqual(len(job.days), 1)
            self.assertEqual(len(job.days[0].fragments), 2)
            self.assertEqual(len(job.files), 2)

    def test_pipeline_selector_is_independent_of_legacy_output_and_display_limit(self) -> None:
        import tempfile

        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as temporary:
            folder = Path(temporary)
            source = folder / "D&TMOK KLNG_20260601_fragment.xlsx"
            self._write_fragment(source, [100.0, 80.0, 60.0, 40.0, 20.0, 19.0, 18.0])
            fragment = SourceFragment("KLNG", dt.date(2026, 6, 1), "FRAGMENT", source)
            day_input = DayInput("KLNG", dt.date(2026, 6, 1), (fragment,))
            common = dict(
                input_dir=folder,
                port="KLNG",
                year=2026,
                month=6,
                top_candidates=3,
            )
            modern_only = process_day_input(
                day_input,
                AppConfig(output_path=folder / "modern.xlsx", **common),
            )
            with_legacy = process_day_input(
                day_input,
                AppConfig(
                    output_path=folder / "both.xlsx",
                    legacy_output_path=folder / "legacy.xlsx",
                    **common,
                ),
            )
            first = modern_only.directions["北"]
            second = with_legacy.directions["北"]
            self.assertEqual(
                (first.selected, first.selected_rank, first.cluster_count),
                (20.0, 5, 3),
            )
            self.assertEqual(
                (second.selected, second.selected_rank, second.cluster_count),
                (first.selected, first.selected_rank, first.cluster_count),
            )

    def test_spool_is_profile_superset_includes_over_cap_and_preserves_provenance(self) -> None:
        import tempfile

        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as temporary:
            folder = Path(temporary)
            source = folder / "D&TMOK KLNG_20260601_any.xlsx"
            self._write_records(
                source,
                [
                    (1, 0, "A", 4, 100.0),
                    (1, 1, "A", 4, 99.0),
                    (1, 2, "B", 4, 98.0),
                    (1, 3, "A", 1, 50.0),
                    (1, 4, "A", 1, 49.0),
                    (1, 5, "B", 1, 48.0),
                    (13, 0, "A", 3, 70.0),
                    (13, 1, "A", 3, 69.0),
                    (13, 2, "B", 3, 68.0),
                    (13, 3, "A", 2, 600.0),
                    (13, 4, "B", 5, 900.0),
                ],
            )
            fragment = SourceFragment("KLNG", dt.date(2026, 6, 1), "ANY", source)
            day_input = DayInput("KLNG", dt.date(2026, 6, 1), (fragment,))
            config = AppConfig(
                input_dir=folder,
                output_path=folder / "modern.xlsx",
                port="KLNG",
                year=2026,
                month=6,
                message_types=(1,),
            )
            result = process_day_input(day_input, config)

            self.assertEqual(result.directions["北"].selected, 50.0)
            self.assertEqual(result.period_directions[PERIOD_A]["北"].selected, 100.0)
            self.assertEqual(result.period_directions[PERIOD_B]["北"].selected, 70.0)
            selected = result.period_directions[PERIOD_A]["北"].candidates[0]
            self.assertTrue(selected.candidate_id)
            self.assertEqual(selected.timestamp, dt.datetime(2026, 6, 1, 1, 0, 0))
            self.assertEqual((selected.msg_type, selected.mmsi, selected.channel), (4, 500000000, "A"))
            self.assertEqual((selected.fragment, selected.sheet, selected.source_row), (source.name, "AIS", 2))

            counts, records = read_normalized_spool(_normalized_spool_path(config, result.day))
            normalized = [record for _direction, record in records]
            self.assertEqual(sum(counts), len(normalized))
            self.assertEqual(len(normalized), 11)
            self.assertEqual({record.msg_type for record in normalized}, {1, 2, 3, 4, 5})
            self.assertEqual(max(record.distance for record in normalized), 900.0)

            alternate = process_day_input(
                day_input,
                AppConfig(
                    input_dir=folder,
                    output_path=folder / "alternate.xlsx",
                    port="KLNG",
                    year=2026,
                    month=6,
                    message_types=(2,),
                ),
            )
            self.assertNotEqual(
                alternate.directions["北"].selected,
                result.directions["北"].selected,
            )
            for period in (PERIOD_A, PERIOD_B):
                self.assertEqual(
                    alternate.period_directions[period]["北"].selected,
                    result.period_directions[period]["北"].selected,
                )

    def test_byte_identical_fragment_alias_is_processed_once(self) -> None:
        import shutil
        import tempfile

        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as temporary:
            folder = Path(temporary)
            first = folder / "D&TMOK KLNG_20260601_copy-a.xlsx"
            second = folder / "D&TMOK KLNG_20260601_copy-b.xlsx"
            self._write_fragment(first, [10.0, 9.9, 9.8])
            shutil.copyfile(first, second)
            catalog, _warnings = scan_source_files(folder)
            config = AppConfig(
                input_dir=folder,
                output_path=folder / "result.xlsx",
                port="KLNG",
                year=2026,
                month=6,
            )
            result = process_day_input(catalog["KLNG"][dt.date(2026, 6, 1)], config)
            self.assertEqual(result.rows_scanned, 3)
            self.assertEqual(result.rows_normalized, 3)
            self.assertEqual(len(result.fragment_info), 1)
            self.assertEqual(result.fragment_info[0].aliases, (second,))
            self.assertTrue(any("byte-identical" in item for item in result.diagnostics))

    def test_official_day_processing_rejects_missing_time_schema(self) -> None:
        import tempfile

        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as temporary:
            folder = Path(temporary)
            source = folder / "D&TMOK KLNG_20260601_bad.xlsx"
            workbook = openpyxl.Workbook(write_only=True)
            sheet = workbook.create_sheet("AIS")
            sheet.append(["msg_type", "LONGITUDE_DESC", "bearing", "distance in nautical miles"])
            sheet.append([1, "East", 0.0, 1.0])
            workbook.save(source)
            fragment = SourceFragment("KLNG", dt.date(2026, 6, 1), "BAD", source)
            config = AppConfig(folder, folder / "result.xlsx", "KLNG", 2026, 6)
            with self.assertRaisesRegex(SourceFileError, "正式五檔輸出需要"):
                process_day_input(DayInput("KLNG", fragment.day, (fragment,)), config)


class ReviewLedgerTests(unittest.TestCase):
    def test_scoped_ledger_enforces_candidate_only_period_contract(self) -> None:
        import shutil
        import tempfile

        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as temporary:
            folder = Path(temporary)
            source = folder / "D&TMOK KLNG_20260601_any.xlsx"
            FragmentCatalogTests._write_records(
                source,
                [
                    (1, 0, "A", 1, 50.0),
                    (1, 1, "A", 1, 49.0),
                    (1, 2, "B", 1, 48.0),
                    (13, 0, "A", 3, 70.0),
                    (13, 1, "A", 3, 69.0),
                    (13, 2, "B", 3, 68.0),
                ],
            )
            fragment = SourceFragment("KLNG", dt.date(2026, 6, 1), "ANY", source)
            config = AppConfig(folder, folder / "analysis.xlsx", "KLNG", 2026, 6)
            result = process_day_input(DayInput("KLNG", fragment.day, (fragment,)), config)
            results = [result]
            snapshot = build_default_decision_snapshot(config, results)
            write_monthly_workbook(config, results, [], snapshot)

            workbook = openpyxl.load_workbook(config.output_path, read_only=False, data_only=False)
            try:
                self.assertIn("決策台帳", workbook.sheetnames)
                self.assertIn("候選清單", workbook.sheetnames)
                self.assertEqual(workbook["系統資料"].sheet_state, "hidden")
                ledger = workbook["決策台帳"]
                self.assertEqual(ledger.max_row, 1 + 3 * 21)
                self.assertEqual(ledger["A2"].value, "logical_day")
                period_a_row = next(
                    row
                    for row in range(2, ledger.max_row + 1)
                    if ledger.cell(row, 1).value == PERIOD_A
                )
                self.assertIsNone(ledger.cell(period_a_row, 7).value)
                self.assertTrue(str(ledger.cell(period_a_row, 8).value).startswith("="))
                candidate_headers = [cell.value for cell in workbook["候選清單"][1]]
                self.assertIn("Timestamp", candidate_headers)
                self.assertIn("MMSI", candidate_headers)
                self.assertIn("Fragment", candidate_headers)
            finally:
                workbook.close()

            candidate_catalog = {
                key: set(candidates)
                for key, candidates in _display_candidate_catalog(results).items()
            }
            loaded = read_review_decisions(
                config.output_path,
                expected_manifest_hash=source_manifest_hash(config, results),
                candidate_catalog=candidate_catalog,
            )
            self.assertEqual(resolve_decision_values(loaded, results), resolve_decision_values(snapshot, results))

            period_numeric = folder / "period_numeric.xlsx"
            shutil.copyfile(config.output_path, period_numeric)
            workbook = openpyxl.load_workbook(period_numeric)
            ledger = workbook["決策台帳"]
            period_a_row = next(
                row for row in range(2, ledger.max_row + 1) if ledger.cell(row, 1).value == PERIOD_A
            )
            ledger.cell(period_a_row, 7).value = 12.3
            workbook.save(period_numeric)
            workbook.close()
            with self.assertRaisesRegex(ValueError, "Period A/B 不允許"):
                read_review_decisions(period_numeric)

            fake_candidate = folder / "fake_candidate.xlsx"
            shutil.copyfile(config.output_path, fake_candidate)
            workbook = openpyxl.load_workbook(fake_candidate)
            workbook["決策台帳"].cell(period_a_row, 6).value = "not-a-real-candidate"
            workbook.save(fake_candidate)
            workbook.close()
            with self.assertRaisesRegex(ValueError, "candidate ID 不屬於"):
                read_review_decisions(fake_candidate, candidate_catalog=candidate_catalog)

            logical_override = folder / "logical_override.xlsx"
            shutil.copyfile(config.output_path, logical_override)
            workbook = openpyxl.load_workbook(logical_override)
            workbook["決策台帳"]["G2"] = 123.456
            workbook.save(logical_override)
            workbook.close()
            overridden = read_review_decisions(logical_override)
            values = resolve_decision_values(overridden, results)
            self.assertEqual(values[("logical_day", dt.date(2026, 6, 1), "北")], 123.456)

            broken_formula = folder / "broken_formula.xlsx"
            shutil.copyfile(config.output_path, broken_formula)
            workbook = openpyxl.load_workbook(broken_formula)
            workbook["決策台帳"]["H2"] = 1
            workbook.save(broken_formula)
            workbook.close()
            with self.assertRaisesRegex(ValueError, "Final value 公式已遭破壞"):
                read_review_decisions(broken_formula)


class ExternalJuneContractTests(unittest.TestCase):
    def test_period_candidate_contract_gate_artifact(self) -> None:
        raw_path = os.environ.get("AIS_JUNE_CONTRACT_GATE_JSON")
        if not raw_path:
            self.skipTest("AIS_JUNE_CONTRACT_GATE_JSON is not configured")
        payload = json.loads(Path(raw_path).read_text(encoding="utf-8"))
        result = payload["results"]
        self.assertEqual(result["nonempty_period_finals"], 1260)
        self.assertEqual(result["matched"], 1260)
        self.assertEqual(result["unmatched"], 0)
        self.assertEqual(result["review_scope_blank_finals"], 0)
        self.assertEqual(result["blank_with_retained_legacy_detail"], 0)
        self.assertEqual(result["structural_land_blank_cells"], 660)
        self.assertEqual(result["structural_land_nonblank_cells"], 0)

    def test_five_file_golden_grid_relationships(self) -> None:
        raw_directory = os.environ.get("AIS_JUNE_GOLDEN_DIR")
        if not raw_directory:
            self.skipTest("AIS_JUNE_GOLDEN_DIR is not configured")
        directory = Path(raw_directory)
        period_a = _read_legacy_grid(directory / "KLNG 6月 32方位數值_2 總表.xlsx")
        period_b = _read_legacy_grid(directory / "KLNG 6月 32方位數值 總表.xlsx")
        integrated = _read_legacy_grid(directory / "KLNG_6月_32方位_每日較大值整合總表.xlsx")

        self.assertEqual(len(period_a), 30 * 32)
        self.assertEqual(len(period_b), 30 * 32)
        self.assertEqual(len(integrated), 30 * 32)
        for key in sorted(integrated):
            left = period_a[key]
            right = period_b[key]
            expected = max(value for value in (left, right) if value is not None) if any(
                value is not None for value in (left, right)
            ) else None
            actual = integrated[key]
            if expected is None:
                self.assertIsNone(actual, key)
            else:
                self.assertTrue(
                    actual is not None
                    and math.isclose(actual, expected, rel_tol=0.0, abs_tol=1e-9),
                    (key, left, right, actual),
                )
        for day in range(1, 31):
            current = dt.date(2026, 6, day)
            for index in range(32):
                if index in OPEN_SEA_INDEXES:
                    self.assertIsNotNone(period_a[(current, index)])
                    self.assertIsNotNone(period_b[(current, index)])
                else:
                    self.assertIsNone(period_a[(current, index)])
                    self.assertIsNone(period_b[(current, index)])


if __name__ == "__main__":
    unittest.main()
