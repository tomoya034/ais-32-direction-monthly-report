from datetime import date
import copy
import hashlib
import json
import math
import os
from pathlib import Path
import struct
import tempfile
import unittest
import zipfile

from scripts.independent_source_audit import (
    CHECKPOINT_SCHEMA, COUNTS, SPOOL_RECORD, RecordCensus, audit_day, audit_reference,
    exact_float_bin, guard_output, independent_identity, resume_results,
    save_report, scan_source, scan_spool, seal_entry,
)


def workbook(path: Path, rows: list[list[object]], title: str = "AIS") -> None:
    def column(index: int) -> str:
        value = ""
        while index:
            index, digit = divmod(index - 1, 26)
            value = chr(65 + digit) + value
        return value

    def cell(value: object, reference: str) -> str:
        if value is None:
            return f'<c r="{reference}"/>'
        if isinstance(value, str):
            from xml.sax.saxutils import escape
            return f'<c r="{reference}" t="inlineStr"><is><t>{escape(value)}</t></is></c>'
        return f'<c r="{reference}"><v>{value}</v></c>'

    xml_rows = [f'<row r="{index}">' + ''.join(cell(value, f'{column(col)}{index}')
                for col, value in enumerate(row, 1)) + '</row>'
                for index, row in enumerate(rows, 1)]
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("xl/workbook.xml", '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="' + title + '" sheetId="1" r:id="rId1"/></sheets></workbook>')
        archive.writestr("xl/_rels/workbook.xml.rels", '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>')
        archive.writestr("xl/worksheets/sheet1.xml", '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + ''.join(xml_rows) + '</sheetData></worksheet>')


HEADERS = ["Year", "Month", "Day", "Hour", "Minute", "Second", "channel",
           "msg_type", "mmsi", "LONGITUDE_DESC", "bearing", "distance in nautical miles"]


class IndependentAuditTests(unittest.TestCase):
    def test_output_and_atomic_temporary_cannot_overwrite_input_bytes(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source_dir, run_dir, report_dir = root / "source", root / "run", root / "reports"
            for directory in (source_dir, run_dir, report_dir):
                directory.mkdir()
            source = source_dir / "raw.xlsx"
            cache = run_dir / "day.json"
            source.write_bytes(b"ORIGINAL_RAW_XLSX")
            cache.write_bytes(b"ORIGINAL_CACHE_JSON")
            for destination in (source, source_dir / "report.json", cache, run_dir / "report.json"):
                with self.assertRaises(ValueError):
                    save_report(destination, {"pass": True}, [source_dir, run_dir], [source, cache])
            alias = report_dir / "raw-alias.json"
            os.link(source, alias)
            with self.assertRaisesRegex(ValueError, "hard-link"):
                save_report(alias, {"pass": True}, [source_dir, run_dir], [source, cache])
            alias.unlink()
            destination = report_dir / "legitimate.json"
            atomic_tmp = destination.with_suffix(".json.tmp")
            os.link(cache, atomic_tmp)
            with self.assertRaisesRegex(ValueError, "hard-link"):
                save_report(destination, {"pass": True}, [source_dir, run_dir], [source, cache])
            self.assertEqual(source.read_bytes(), b"ORIGINAL_RAW_XLSX")
            self.assertEqual(cache.read_bytes(), b"ORIGINAL_CACHE_JSON")
            self.assertFalse(destination.exists())

    def test_reference_parent_tree_is_read_only(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            reference_dir = root / "reference"
            reference_dir.mkdir()
            reference = reference_dir / "manual.xlsx"
            reference.write_bytes(b"CORRECT_REFERENCE")
            with self.assertRaises(ValueError):
                save_report(reference, {}, [reference_dir], [reference])
            with self.assertRaises(ValueError):
                save_report(reference_dir / "report.json", {}, [reference_dir], [reference])
            save_report(root / "legitimate.json", {"pass": True}, [reference_dir], [reference])
            self.assertEqual(reference.read_bytes(), b"CORRECT_REFERENCE")

    def test_resume_rejects_unsealed_or_tampered_pass_and_foreign_cache(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "raw.xlsx"
            cache = root / "day.json"
            spool = root / "day.bin"
            source.write_bytes(b"raw")
            spool.write_bytes(b"spool")
            counters = {name: 0 for name in ("rows_scanned", "rows_accepted", "rows_invalid",
                                           "rows_wrong_message", "rows_not_east", "rows_legacy", "rows_normalized")}
            cache.write_text(json.dumps({"result": counters}), encoding="utf-8")
            task = {"port": "HWLN", "day": "2026-04-01", "paths": [str(source)],
                    "cache": str(cache), "spool": str(spool)}
            census = RecordCensus().evidence()
            entry = seal_entry({"port": "HWLN", "day": "2026-04-01", "issues": [],
                                "pass": True, "source_files": 1, "cache": str(cache), "spool": str(spool),
                                "totals": counters, "raw_census": census, "spool_census": census}, task)
            prior = {"checkpoint_schema": CHECKPOINT_SCHEMA, "source_root": str(root.resolve()),
                     "run_root": str(root.resolve()), "results": [entry]}
            self.assertEqual(len(resume_results(prior, [task], root, root)), 1)
            unsealed = copy.deepcopy(prior)
            unsealed.pop("checkpoint_schema")
            with self.assertRaisesRegex(ValueError, "unsealed"):
                resume_results(unsealed, [task], root, root)
            tampered = copy.deepcopy(prior)
            tampered["results"][0]["pass"] = False
            with self.assertRaisesRegex(ValueError, "tampered"):
                resume_results(tampered, [task], root, root)
            forged_pass = copy.deepcopy(prior)
            forged_pass["results"][0]["issues"] = ["unresolved failure"]
            forged_pass["results"][0].pop("entry_sha256")
            forged_pass["results"][0] = seal_entry(forged_pass["results"][0], task)
            with self.assertRaisesRegex(ValueError, "retains issues"):
                resume_results(forged_pass, [task], root, root)
            foreign = copy.deepcopy(prior)
            foreign["results"][0]["cache"] = str(root / "outside-day.json")
            foreign["results"][0].pop("entry_sha256")
            foreign["results"][0] = seal_entry(foreign["results"][0], task)
            with self.assertRaisesRegex(ValueError, "outside planned"):
                resume_results(foreign, [task], root, root)
            source.write_bytes(b"changed raw")
            with self.assertRaisesRegex(ValueError, "inputs changed"):
                resume_results(prior, [task], root, root)

    def test_filename_date_is_independent_and_unambiguous(self):
        self.assertEqual(independent_identity("D&TMOK HWLN_2026041_23.xlsx"), ("HWLN", date(2026, 4, 1)))
        self.assertEqual(independent_identity("D&TMOK HWLN_20260401_23.xlsx"), ("HWLN", date(2026, 4, 1)))
        with self.assertRaises(ValueError):
            independent_identity("D&TMOK HWLN_20260230_23.xlsx")

    def test_every_exact_boundary_and_adjacent_float(self):
        for index in range(32):
            boundary = index * 11.25
            self.assertEqual(exact_float_bin(boundary), index)
            if index:
                self.assertEqual(exact_float_bin(math.nextafter(boundary, -math.inf)), index - 1)
            self.assertEqual(exact_float_bin(math.nextafter(boundary, math.inf)), index)
        self.assertEqual(exact_float_bin(360), 0)
        self.assertEqual(exact_float_bin(-11.25), 31)
        self.assertEqual(exact_float_bin(-math.ulp(0.0)), 31)
        self.assertIsNone(exact_float_bin(float("nan")))

    def test_source_counters_filter_order_and_period_boundaries(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "D&TMOK HWLN_2026041_23.xlsx"
            rows = [HEADERS,
                    [2026, 4, 1, 11, 59, 59, "A", 1, 123, "East", 11.25, 20],
                    [2026, 4, 1, 12, 0, 0, "A", 4, 123, "East", 359.9, 600],
                    [2026, 4, 1, 12, 0, 1, "B", 2, 123, "West", 0, 2],
                    [2026, 4, 1, 12, 0, 2, "B", 99, 123, "East", 45, -1],
                    [2026, 4, 1, 12, 0, 3, "B", 1.5, 123, "West", 45, 2]]
            workbook(path, rows)
            census = RecordCensus()
            result = scan_source(path, date(2026, 4, 1), {1, 2, 3, 18, 19}, 500, census)
            self.assertEqual([result[name] for name in (
                "rows_scanned", "rows_accepted", "rows_invalid", "rows_wrong_message",
                "rows_not_east", "rows_legacy", "rows_normalized")], [5, 1, 2, 2, 1, 2, 2])
            self.assertEqual(census.histogram[(1, 0, 1)], 1)
            self.assertEqual(census.histogram[(31, 1, 4)], 1)

    def test_wrong_day_is_a_hard_failure(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "raw.xlsx"
            workbook(path, [HEADERS, [2026, 4, 2, 0, 0, 0, "A", 1, 123, "East", 0, 2]])
            with self.assertRaisesRegex(ValueError, "!="):
                scan_source(path, date(2026, 4, 1), {1}, 500, RecordCensus())

    def test_raw_spool_multiset_detects_value_change_with_same_count(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "spool.bin"
            counts = [0] * 32
            counts[1] = 1
            path.write_bytes(b"AISNRM1\0" + COUNTS.pack(*counts) +
                             SPOOL_RECORD.pack(20, 11.25, 2, 43199, 1, 0, 123, b"A".ljust(16, b"\0")))
            spool = scan_spool(path)
            raw = RecordCensus()
            raw.add(1, 20, 11.25, 43199, 1, 123, "A", source_key=2)
            self.assertEqual(raw.evidence(), spool.evidence())
            changed = RecordCensus()
            changed.add(1, 21, 11.25, 43199, 1, 123, "A", source_key=2)
            self.assertEqual(changed.counts, spool.counts)
            self.assertNotEqual(changed.evidence(), spool.evidence())

    def test_spool_period_corruption_is_rejected(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "spool.bin"
            counts = [0] * 32
            counts[0] = 1
            path.write_bytes(b"AISNRM1\0" + COUNTS.pack(*counts) +
                             SPOOL_RECORD.pack(20, 0, 2, 43200, 1, 0, 123, b"A".ljust(16, b"\0")))
            with self.assertRaisesRegex(ValueError, "period"):
                scan_spool(path)

    def test_identical_values_cannot_hide_wrong_source_row_or_fragment(self):
        correct = RecordCensus()
        correct.add(0, 20, 0, 0, 1, 123, "A", source_key=2)
        for wrong_key in (3, (1 << 32) | 2):
            wrong = RecordCensus()
            wrong.add(0, 20, 0, 0, 1, 123, "A", source_key=wrong_key)
            self.assertEqual(correct.counts, wrong.counts)
            self.assertEqual(correct.histogram, wrong.histogram)
            self.assertNotEqual(correct.evidence(), wrong.evidence())
        for wrong_key in (0, 1, 1 << 32):
            with self.assertRaisesRegex(ValueError, "source key"):
                RecordCensus().add(0, 20, 0, 0, 1, 123, "A", source_key=wrong_key)
        with self.assertRaisesRegex(ValueError, "MMSI"):
            RecordCensus().add(0, 20, 0, 0, 1, -2, "A", source_key=2)

    def test_full_raw_provenance_rejects_equal_value_wrong_keys_and_fragment_metadata(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            paths = [root / "D&TMOK KLNG_20260601_11.xlsx", root / "D&TMOK KLNG_20260601_23.xlsx"]
            for second, path in enumerate(paths):
                workbook(path, [HEADERS, [2026, 6, 1, 0, 0, second, "A", 1, 123, "East", 0, 20]])
            fingerprints = sorted([(hashlib.sha256(path.read_bytes()).hexdigest().upper(), path)
                                   for path in paths])
            fragment_info = []
            records = []
            for index, (digest, path) in enumerate(fingerprints):
                second = paths.index(path)
                fragment_info.append({"path": str(path), "sha256": digest, "sheet": "AIS",
                                      "rows_scanned": 1, "min_second": second, "max_second": second,
                                      "occupied_seconds": 1})
                records.append(SPOOL_RECORD.pack(20, 0, (index << 32) | 2, second, 1, 0, 123,
                                                 b"A".ljust(16, b"\0")))
            counts = [2] + [0] * 31
            spool = root / "20260601.bin"
            spool.write_bytes(b"AISNRM1\0" + COUNTS.pack(*counts) + b"".join(records))
            cache = root / "20260601.json"
            payload = {"signature": {"port": "KLNG", "date": "2026-06-01",
                                     "sources": [{"path": str(path)} for path in paths],
                                     "message_types": [1, 2, 3, 18, 19], "max_distance": 500,
                                     "direction_policy": "klng21"},
                       "result": {"day": "2026-06-01", "rows_scanned": 2, "rows_accepted": 2,
                                  "rows_invalid": 0, "rows_wrong_message": 0, "rows_not_east": 0,
                                  "rows_legacy": 2, "rows_normalized": 2, "fragment_info": fragment_info}}
            cache.write_text(json.dumps(payload), encoding="utf-8")
            task = {"port": "KLNG", "day": "2026-06-01", "paths": [str(path) for path in paths],
                    "cache": str(cache), "spool": str(spool)}
            self.assertTrue(audit_day(task)["pass"])
            first = list(SPOOL_RECORD.unpack(records[0]))
            first[2] += 1  # same numeric AIS record but a different original Excel row
            spool.write_bytes(b"AISNRM1\0" + COUNTS.pack(*counts) + SPOOL_RECORD.pack(*first) + records[1])
            self.assertFalse(audit_day(task)["pass"])
            spool.write_bytes(b"AISNRM1\0" + COUNTS.pack(*counts) + b"".join(records))
            for mutation in ("order", "hash", "sheet", "signature_path"):
                bad = copy.deepcopy(payload)
                if mutation == "order":
                    bad["result"]["fragment_info"].reverse()
                elif mutation == "hash":
                    bad["result"]["fragment_info"][0]["sha256"] = "0" * 64
                elif mutation == "sheet":
                    bad["result"]["fragment_info"][0]["sheet"] = "wrong sheet"
                else:
                    bad["signature"]["sources"][0]["path"] = str(root / "other" / paths[0].name)
                cache.write_text(json.dumps(bad), encoding="utf-8")
                self.assertFalse(audit_day(task)["pass"], mutation)

    def test_v2_value_only_census_is_not_a_source_identity_pass(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            prior = {"checkpoint_schema": "AIS_INDEPENDENT_SOURCE_AUDIT_2", "results": []}
            with self.assertRaisesRegex(ValueError, "unsealed"):
                resume_results(prior, [], root, root)

    def test_reference_processes_every_detail_row_and_legacy_label(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "reference.xlsx"
            workbook(path, [["distance", "bearing", "direction"],
                            [1, 22.5, "北北東"], [2, 33.75, "東北微北"],
                            [3, 359.99, "北微西"]], title="6月1日")
            result = audit_reference(path)
            self.assertEqual(result["detail_rows"], 3)
            self.assertEqual(result["matched_rows"], 3)
            self.assertTrue(result["pass"])

    def test_audit_day_compares_manual_counter_oracle_and_value_fingerprint(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "D&TMOK HWLN_2026041_23.xlsx"
            workbook(source, [HEADERS, [2026, 4, 1, 12, 0, 0, "A", 1, 123, "East", 11.25, 20]])
            spool = root / "20260401.bin"
            counts = [0] * 32
            counts[1] = 1
            spool.write_bytes(b"AISNRM1\0" + COUNTS.pack(*counts) +
                              SPOOL_RECORD.pack(20, 11.25, 2, 43200, 1, 1, 123, b"A".ljust(16, b"\0")))
            cache = root / "20260401.json"
            payload = {"signature": {"port": "HWLN", "date": "2026-04-01", "sources": [{"path": str(source)}],
                                     "message_types": [1, 2, 3, 18, 19], "max_distance": 500,
                                     "direction_policy": "all32"},
                       "result": {"day": "2026-04-01", "rows_scanned": 1, "rows_accepted": 1,
                                  "rows_invalid": 0, "rows_wrong_message": 0, "rows_not_east": 0,
                                  "rows_legacy": 1, "rows_normalized": 1,
                                  "fragment_info": [{"path": str(source), "sha256": hashlib.sha256(source.read_bytes()).hexdigest(), "sheet": "AIS", "rows_scanned": 1,
                                                     "min_second": 43200, "max_second": 43200, "occupied_seconds": 1}]}}
            cache.write_text(json.dumps(payload), encoding="utf-8")
            task = {"port": "HWLN", "day": "2026-04-01", "paths": [str(source)],
                    "cache": str(cache), "spool": str(spool)}
            self.assertTrue(audit_day(task)["pass"])
            payload["result"]["rows_scanned"] = 2
            cache.write_text(json.dumps(payload), encoding="utf-8")
            self.assertFalse(audit_day(task)["pass"])
            payload["result"]["rows_scanned"] = 1
            cache.write_text(json.dumps(payload), encoding="utf-8")
            spool.write_bytes(b"AISNRM1\0" + COUNTS.pack(*counts) +
                              SPOOL_RECORD.pack(21, 11.25, 2, 43200, 1, 1, 123, b"A".ljust(16, b"\0")))
            self.assertIn("raw/spool normalized record multiset/count mismatch", audit_day(task)["issues"])


if __name__ == "__main__":
    unittest.main()
