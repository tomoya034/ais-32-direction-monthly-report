from __future__ import annotations

import datetime as dt
import json
import math
import os
import unittest
from pathlib import Path
from unittest import mock

import openpyxl
from test_direction_contracts import V150_DIRECTION_ORDER

from ais_monthly_app import (
    APP_VERSION,
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
    derive_delivery_paths,
    finalize_review_workbook,
    parse_source_filename,
    process_day_input,
    process_month,
    read_normalized_spool,
    read_review_decisions,
    resolve_decision_values,
    scan_source_files,
    select_cluster_from_sorted_rows,
    source_manifest_hash,
    write_delivery_workbooks,
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


class ReleaseMetadataTests(unittest.TestCase):
    def test_v15_version_metadata_are_consistent(self) -> None:
        import tomllib

        root = Path(__file__).parent
        version = (root / "VERSION").read_text(encoding="utf-8").strip()
        project = tomllib.loads((root / "pyproject.toml").read_text(encoding="utf-8"))
        version_info = (root / "version_info.txt").read_text(encoding="utf-8")
        readme = (root / "README.md").read_text(encoding="utf-8")
        usage = (root / "使用說明.txt").read_text(encoding="utf-8")

        self.assertEqual(APP_VERSION, "1.5.1")
        self.assertEqual(version, APP_VERSION)
        self.assertEqual(project["project"]["version"], APP_VERSION)
        self.assertIn("filevers=(1, 5, 1, 0)", version_info)
        self.assertIn("prodvers=(1, 5, 1, 0)", version_info)
        self.assertIn("v1.5.1", readme)
        self.assertIn("v1.5.1", usage)


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

    def test_max_days_never_cuts_a_logical_day_fragment_set(self) -> None:
        import tempfile

        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as temporary:
            folder = Path(temporary)
            first = folder / "D&TMOK KLNG_20260601_11.xlsx"
            second = folder / "D&TMOK KLNG_20260601_23.xlsx"
            later = folder / "D&TMOK KLNG_20260602_11.xlsx"
            self._write_fragment(first, [100.0, 99.0])
            self._write_fragment(second, [98.0])
            self._write_fragment(later, [70.0, 69.0, 68.0])
            config = AppConfig(
                folder,
                folder / "analysis.xlsx",
                "KLNG",
                2026,
                6,
                top_candidates=3,
                max_days=1,
            )
            results, _warnings = process_month(config)
            self.assertEqual(len(results[0].source_fragments), 2)
            self.assertEqual(results[0].directions["北"].selected, 100.0)
            self.assertFalse(results[1].source_fragments)
            self.assertIn("logical day", results[1].note)

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

    def test_time_range_overlap_without_occupied_second_short_circuits_duplicate_scan(self) -> None:
        import tempfile

        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as temporary:
            folder = Path(temporary)
            first = folder / "D&TMOK KLNG_20260601_part-a.xlsx"
            second = folder / "D&TMOK KLNG_20260601_part-b.xlsx"
            self._write_records(first, [(1, 0, "A", 1, 50.0), (1, 2, "A", 1, 48.0)])
            self._write_records(second, [(1, 1, "B", 3, 49.0)])
            catalog, _warnings = scan_source_files(folder)
            config = AppConfig(folder, folder / "result.xlsx", "KLNG", 2026, 6)
            result = process_day_input(catalog["KLNG"][dt.date(2026, 6, 1)], config)
            joined = "\n".join(result.diagnostics)
            self.assertIn("time-range overlap", joined)
            self.assertNotIn("occupied-second overlap", joined)
            self.assertNotIn("possible duplicate events", joined)
            self.assertEqual(result.rows_normalized, 3)

    def test_occupied_overlap_reports_normal_possible_and_exact_without_dedupe(self) -> None:
        import tempfile

        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as temporary:
            folder = Path(temporary)
            first = folder / "D&TMOK KLNG_20260601_alpha.xlsx"
            second = folder / "D&TMOK KLNG_20260601_beta.xlsx"
            self._write_records(
                first,
                [
                    (1, 0, "A", 1, 50.0),
                    (1, 1, "A", 3, 40.0),
                    (1, 2, "A", 1, 30.0),
                ],
            )
            self._write_records(
                second,
                [
                    (1, 0, "B", 1, 49.0),
                    (1, 1, "A", 3, 40.0),
                    (1, 2, "A", 3, 29.0),
                ],
            )
            catalog, _warnings = scan_source_files(folder)
            config = AppConfig(folder, folder / "result.xlsx", "KLNG", 2026, 6)
            result = process_day_input(catalog["KLNG"][dt.date(2026, 6, 1)], config)
            joined = "\n".join(result.diagnostics)
            self.assertIn("occupied-second overlap", joined)
            self.assertIn("normal same-second different AIS messages：1 秒", joined)
            self.assertIn("possible duplicate events：1 筆", joined)
            self.assertIn("exact normalized core records：1 筆", joined)
            self.assertIn("未自動去重", joined)
            self.assertEqual(result.rows_normalized, 6)

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

    def test_official_day_processing_rejects_row_from_another_day(self) -> None:
        import tempfile

        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as temporary:
            folder = Path(temporary)
            source = folder / "D&TMOK KLNG_20260601_wrong-day.xlsx"
            workbook = openpyxl.Workbook(write_only=True)
            sheet = workbook.create_sheet("AIS")
            sheet.append([
                "Year", "Month", "Day", "Hour", "Minute", "Second", "channel",
                "msg_type", "mmsi", "LONGITUDE_DESC", "bearing",
                "distance in nautical miles",
            ])
            sheet.append([2026, 6, 2, 1, 0, 0, "A", 1, 500000001, "East", 0.0, 50.0])
            workbook.save(source)
            fragment = SourceFragment("KLNG", dt.date(2026, 6, 1), "WRONG-DAY", source)
            config = AppConfig(folder, folder / "result.xlsx", "KLNG", 2026, 6)
            with self.assertRaisesRegex(SourceFileError, "不符合檔名 logical day"):
                process_day_input(DayInput("KLNG", fragment.day, (fragment,)), config)

    def test_official_day_processing_rejects_corrupt_xlsx(self) -> None:
        import tempfile

        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as temporary:
            folder = Path(temporary)
            source = folder / "D&TMOK KLNG_20260601_corrupt.xlsx"
            source.write_bytes(b"not an xlsx")
            fragment = SourceFragment("KLNG", dt.date(2026, 6, 1), "CORRUPT", source)
            config = AppConfig(folder, folder / "result.xlsx", "KLNG", 2026, 6)
            with self.assertRaisesRegex(SourceFileError, "無法開啟來源檔"):
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
            workbook["決策台帳"]["H2"] = "=1"
            workbook.save(broken_formula)
            workbook.close()
            with self.assertRaisesRegex(ValueError, "Final value 公式已遭破壞"):
                read_review_decisions(broken_formula)

    def test_five_delivery_workbooks_share_one_snapshot_and_keep_max_contracts(self) -> None:
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
                    (13, 3, "A", 2, 900.0),
                ],
            )
            fragment = SourceFragment("KLNG", dt.date(2026, 6, 1), "ANY", source)
            config = AppConfig(
                folder,
                folder / "analysis.xlsx",
                "KLNG",
                2026,
                6,
                overwrite=True,
            )
            result = process_day_input(DayInput("KLNG", fragment.day, (fragment,)), config)
            snapshot = build_default_decision_snapshot(config, [result])
            paths = write_delivery_workbooks(
                config,
                [result],
                snapshot,
                root=folder / "delivery",
            )
            self.assertEqual(len(paths.all_files()), 5)
            self.assertTrue(all(path.is_file() for path in paths.all_files()))

            def cell(path: Path, sheet: str, coordinate: str):
                workbook = openpyxl.load_workbook(path, read_only=False, data_only=True)
                try:
                    return workbook[sheet][coordinate].value
                finally:
                    workbook.close()

            self.assertEqual(cell(paths.period_a_full, "總表", "B2"), 50.0)
            self.assertEqual(cell(paths.period_a_summary, "總表", "B2"), 50.0)
            self.assertEqual(cell(paths.period_b_full, "總表", "B2"), 70.0)
            self.assertEqual(cell(paths.period_b_summary, "總表", "B2"), 70.0)
            self.assertEqual(cell(paths.integrated_summary, "總表", "B2"), 70.0)

            for full_path, summary_path in (
                (paths.period_a_full, paths.period_a_summary),
                (paths.period_b_full, paths.period_b_summary),
            ):
                full = openpyxl.load_workbook(full_path, read_only=True, data_only=True)
                small = openpyxl.load_workbook(summary_path, read_only=True, data_only=True)
                try:
                    full_row = tuple(
                        full["總表"].iter_rows(
                            min_row=2, max_row=2, min_col=2, max_col=33, values_only=True
                        )
                    )[0]
                    small_row = tuple(
                        small["總表"].iter_rows(
                            min_row=2, max_row=2, min_col=2, max_col=33, values_only=True
                        )
                    )[0]
                    self.assertEqual(full_row, small_row)
                finally:
                    full.close()
                    small.close()

            workbook = openpyxl.load_workbook(paths.period_a_full, read_only=True, data_only=True)
            try:
                retained = [
                    row[0]
                    for row in workbook["6月1日"].iter_rows(min_row=2, max_col=3, values_only=True)
                    if row[2] == "北"
                ]
                self.assertEqual(max(retained), cell(paths.period_a_full, "6月1日", "H2"))
                self.assertNotIn(900.0, retained)
            finally:
                workbook.close()

            formulas = openpyxl.load_workbook(paths.integrated_summary, read_only=False, data_only=False)
            try:
                self.assertIn("MAX('Period A'!B2,'Period B'!B2)", formulas["總表"]["B2"].value)
                self.assertEqual(len(formulas["總表"]._charts), 2)
            finally:
                formulas.close()
            for path in paths.all_files():
                formulas = openpyxl.load_workbook(path, read_only=False, data_only=False)
                try:
                    summary = formulas["總表"]
                    self.assertEqual(len(summary._charts), 2)
                    self.assertIn("MAX(B2:B2)", summary["B4"].value)
                    for row in summary.iter_rows():
                        for item in row:
                            if isinstance(item.value, str) and item.value.startswith("="):
                                self.assertNotIn("#REF!", item.value)
                finally:
                    formulas.close()

            decisions = list(snapshot.decisions)
            for index, decision in enumerate(decisions):
                if (
                    decision.scope == PERIOD_A
                    and decision.day == dt.date(2026, 6, 1)
                    and decision.direction == "北"
                ):
                    decisions[index] = type(decision)(
                        decision.scope,
                        decision.day,
                        decision.direction,
                        None,
                        None,
                        "blank means suppress retained detail",
                    )
                    break
            blank_snapshot = type(snapshot)(
                snapshot.port,
                snapshot.year,
                snapshot.month,
                snapshot.source_manifest_hash,
                tuple(decisions),
            )
            blank_paths = write_delivery_workbooks(
                config,
                [result],
                blank_snapshot,
                root=folder / "blank_delivery",
            )
            self.assertIsNone(cell(blank_paths.period_a_summary, "總表", "B2"))
            self.assertEqual(cell(blank_paths.integrated_summary, "總表", "B2"), 70.0)
            workbook = openpyxl.load_workbook(blank_paths.period_a_full, read_only=True, data_only=True)
            try:
                self.assertFalse(
                    any(
                        row[2] == "北"
                        for row in workbook["6月1日"].iter_rows(
                            min_row=2, max_col=3, values_only=True
                        )
                    )
                )
            finally:
                workbook.close()

    def test_finalize_review_validates_sources_and_regenerates_all_five(self) -> None:
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
            analysis = folder / "analysis.xlsx"
            config = AppConfig(
                folder,
                analysis,
                "KLNG",
                2026,
                6,
                overwrite=True,
            )
            day_results, warnings = process_month(config)
            snapshot = build_default_decision_snapshot(config, day_results)
            write_monthly_workbook(config, day_results, warnings, snapshot)
            june_first = next(result for result in day_results if result.day.day == 1)
            replacement = next(
                candidate
                for candidate in june_first.period_directions[PERIOD_A]["北"].candidates
                if math.isclose(candidate.distance, 49.0)
            )

            workbook = openpyxl.load_workbook(analysis)
            ledger = workbook["決策台帳"]
            period_a_row = next(
                row
                for row in range(2, ledger.max_row + 1)
                if ledger.cell(row, 1).value == PERIOD_A
                and ledger.cell(row, 2).value.day == 1
                and ledger.cell(row, 3).value == "北"
            )
            logical_row = next(
                row
                for row in range(2, ledger.max_row + 1)
                if ledger.cell(row, 1).value == "logical_day"
                and ledger.cell(row, 2).value.day == 1
                and ledger.cell(row, 3).value == "北"
            )
            ledger.cell(period_a_row, 6).value = replacement.candidate_id
            ledger.cell(logical_row, 7).value = 999.0
            workbook.save(analysis)
            workbook.close()

            paths = finalize_review_workbook(
                analysis,
                delivery_dir=folder / "finalized",
                overwrite=True,
            )
            self.assertTrue(all(path.is_file() for path in paths.all_files()))

            def cell(path: Path, coordinate: str):
                result = openpyxl.load_workbook(path, read_only=True, data_only=True)
                try:
                    return result["總表"][coordinate].value
                finally:
                    result.close()

            self.assertEqual(cell(paths.period_a_full, "B2"), 49.0)
            self.assertEqual(cell(paths.period_a_summary, "B2"), 49.0)
            self.assertEqual(cell(paths.period_b_full, "B2"), 70.0)
            self.assertEqual(cell(paths.period_b_summary, "B2"), 70.0)
            self.assertEqual(cell(paths.integrated_summary, "B2"), 70.0)

            fake = folder / "fake_review.xlsx"
            shutil.copyfile(analysis, fake)
            workbook = openpyxl.load_workbook(fake)
            workbook["決策台帳"].cell(period_a_row, 6).value = "not-a-source-candidate"
            workbook.save(fake)
            workbook.close()
            with self.assertRaisesRegex(ValueError, "實際來源候選"):
                finalize_review_workbook(
                    fake,
                    delivery_dir=folder / "fake_delivery",
                    overwrite=True,
                )

            incomplete = folder / "incomplete_review.xlsx"
            shutil.copyfile(analysis, incomplete)
            workbook = openpyxl.load_workbook(incomplete)
            workbook["決策台帳"].delete_rows(workbook["決策台帳"].max_row)
            workbook.save(incomplete)
            workbook.close()
            with self.assertRaisesRegex(ValueError, "列數已變更"):
                finalize_review_workbook(
                    incomplete,
                    delivery_dir=folder / "incomplete_delivery",
                    overwrite=True,
                )

            stat = source.stat()
            os.utime(source, ns=(stat.st_atime_ns, stat.st_mtime_ns + 1_000_000_000))
            with self.assertRaisesRegex(ValueError, "已變更"):
                finalize_review_workbook(
                    analysis,
                    delivery_dir=folder / "changed_source",
                    overwrite=True,
                )

    def test_delivery_row_limit_errors_before_any_final_file_is_published(self) -> None:
        import tempfile

        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as temporary:
            folder = Path(temporary)
            source = folder / "D&TMOK KLNG_20260601_any.xlsx"
            FragmentCatalogTests._write_records(
                source,
                [
                    (13, 0, "A", 1, 70.0),
                    (13, 1, "A", 1, 69.0),
                    (13, 2, "B", 1, 68.0),
                ],
            )
            fragment = SourceFragment("KLNG", dt.date(2026, 6, 1), "ANY", source)
            config = AppConfig(
                folder,
                folder / "analysis.xlsx",
                "KLNG",
                2026,
                6,
                delivery_dir=folder / "delivery",
                overwrite=True,
            )
            result = process_day_input(DayInput("KLNG", fragment.day, (fragment,)), config)
            snapshot = build_default_decision_snapshot(config, [result])
            expected = derive_delivery_paths(config)
            with mock.patch("ais_monthly_app.EXCEL_MAX_DATA_ROWS", 2):
                with self.assertRaisesRegex(SourceFileError, "超過 Excel 上限"):
                    write_delivery_workbooks(config, [result], snapshot)
            self.assertFalse(any(path.exists() for path in expected.all_files()))
            self.assertFalse(any(expected.root.glob("*.building.xlsx")))

    def test_five_file_promotion_failure_restores_previous_snapshot(self) -> None:
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
            config = AppConfig(
                folder,
                folder / "analysis.xlsx",
                "KLNG",
                2026,
                6,
                delivery_dir=folder / "delivery",
                overwrite=True,
            )
            result = process_day_input(DayInput("KLNG", fragment.day, (fragment,)), config)
            snapshot = build_default_decision_snapshot(config, [result])
            paths = derive_delivery_paths(config)
            paths.root.mkdir(parents=True, exist_ok=True)
            old_contents = {}
            for index, path in enumerate(paths.all_files()):
                content = f"old-snapshot-{index}".encode("ascii")
                path.write_bytes(content)
                old_contents[path] = content

            real_replace = os.replace
            building_promotions = 0

            def fail_second_building_promotion(source_path, destination_path):
                nonlocal building_promotions
                if str(source_path).endswith(".building.xlsx"):
                    building_promotions += 1
                    if building_promotions == 2:
                        raise OSError("simulated promotion failure")
                return real_replace(source_path, destination_path)

            with mock.patch("ais_monthly_app.os.replace", side_effect=fail_second_building_promotion):
                with self.assertRaisesRegex(OSError, "simulated promotion failure"):
                    write_delivery_workbooks(config, [result], snapshot)

            for path, content in old_contents.items():
                self.assertEqual(path.read_bytes(), content)
            self.assertFalse(any(paths.root.glob("*.building.xlsx")))
            self.assertFalse(any(paths.root.glob("*.rollback")))


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

    def test_representative_days_match_external_union_baseline(self) -> None:
        import hashlib
        import struct
        import tempfile

        input_path = os.environ.get("AIS_JUNE_INPUT_DIR")
        baseline_path = os.environ.get("AIS_JUNE_BASELINE_JSON")
        benchmark_path = os.environ.get("AIS_JUNE_BENCHMARK_JSON")
        if not input_path or not baseline_path or not benchmark_path:
            self.skipTest("external June representative-day regression is not configured")
        input_dir = Path(input_path)
        baseline = json.loads(Path(baseline_path).read_text(encoding="utf-8"))
        benchmark = json.loads(Path(benchmark_path).read_text(encoding="utf-8"))
        baseline_by_day = {item["date"]: item for item in baseline["days"]}
        benchmark_by_day = {
            item["date"].replace("-", ""): item for item in benchmark["days"]
        }
        selected_days = tuple(
            item.strip()
            for item in os.environ.get(
                "AIS_JUNE_REGRESSION_DAYS", "20260601,20260620"
            ).split(",")
            if item.strip()
        )
        output_parent = Path(
            os.environ.get("AIS_JUNE_TEST_OUTPUT_ROOT", tempfile.gettempdir())
        )
        output_parent.mkdir(parents=True, exist_ok=True)

        def selector_checksum(result) -> str:
            digest = hashlib.sha256()
            for index, direction in enumerate(DIRECTION_ORDER):
                selected = result.directions[direction]
                digest.update(struct.pack("<B", index))
                digest.update(
                    struct.pack(
                        "<d",
                        float("nan") if selected.selected is None else selected.selected,
                    )
                )
                digest.update(
                    struct.pack(
                        "<q", -1 if selected.selected_rank is None else selected.selected_rank
                    )
                )
                digest.update(struct.pack("<q", selected.cluster_count))
            return digest.hexdigest()

        with tempfile.TemporaryDirectory(dir=output_parent) as temporary:
            config = AppConfig(
                input_dir,
                Path(temporary) / "analysis.xlsx",
                "KLNG",
                2026,
                6,
                workers=1,
                overwrite=True,
            )
            job = build_processing_job(config)
            job_by_day = {item.day.strftime("%Y%m%d"): item for item in job.days}
            for day_key in selected_days:
                with self.subTest(day=day_key):
                    result = process_day_input(job_by_day[day_key], config)
                    expected = baseline_by_day[day_key]
                    measured = benchmark_by_day[day_key]["simple"]
                    self.assertEqual(result.rows_scanned, benchmark_by_day[day_key]["raw_rows"])
                    self.assertEqual(result.rows_accepted, measured["eligible_under_cap_rows"])
                    self.assertEqual(len(result.source_fragments), 2)
                    expected_by_index = {
                        index: expected["current_full"][old_label]
                        for index, old_label in enumerate(V150_DIRECTION_ORDER)
                    }
                    for index, direction in enumerate(DIRECTION_ORDER):
                        actual_direction = result.directions[direction]
                        expected_direction = expected_by_index[index]
                        self.assertEqual(
                            (
                                actual_direction.selected,
                                actual_direction.selected_rank,
                                actual_direction.cluster_count,
                            ),
                            (
                                expected_direction["selected"],
                                expected_direction["rank"],
                                expected_direction["cluster_count"],
                            ),
                        )
                    self.assertEqual(
                        selector_checksum(result), measured["selector_sha256"]
                    )


if __name__ == "__main__":
    unittest.main()
