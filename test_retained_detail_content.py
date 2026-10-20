import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest import mock

import xlsxwriter
from scripts import validate_full_regression as audit


class RetainedDetailContentTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="ais-retained-oracle-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.spool = self.root / "rows"
        self.spool.mkdir()
        self.day = date(2026, 6, 1)
        self.path = self.root / "HWLN 6月 32方位數值_2.xlsx"

    def source(self, records):
        groups = [[] for _ in range(32)]
        for distance, bearing, second, msg in records:
            groups[audit.independent_detail_bin(bearing)].append((distance, bearing, second, msg))
        with (self.spool / "20260601.bin").open("wb") as handle:
            handle.write(b"AISNRM1\0")
            handle.write(audit.DETAIL_COUNTS.pack(*(len(group) for group in groups)))
            for group in groups:
                for key, (distance, bearing, second, msg) in enumerate(group, 2):
                    handle.write(audit.DETAIL_ROW.pack(distance, bearing, key, second, msg,
                                                     int(second >= 43200), -1, b"A"))

    def workbook(self, records, maximum=5):
        workbook = xlsxwriter.Workbook(self.path)
        sheet = workbook.add_worksheet("6月1日")
        sheet.write_row(0, 0, ["DISTANCE", "DEGREE", "方位"])
        for row, values in enumerate(records, 1):
            sheet.write_row(row, 0, values)
        summary = workbook.add_worksheet("總表")
        summary.write_number(1, 1, maximum)
        workbook.close()

    def inspect(self):
        return audit.inspect_detail_mapping(self.path, [self.day], {0}, self.spool)

    def standard(self, rows=None):
        self.source([(5, 0, 0, 1), (4, 0, 1, 1), (3, 0, 2, 1)])
        self.workbook([(5, 0, "北"), (4, 0, "北"), (3, 0, "北")] if rows is None else rows)

    def test_every_real_row_matches(self):
        self.standard()
        result = self.inspect()
        self.assertEqual(result["detail_rows"], 3)
        self.assertEqual(result["retained_expected_rows"], 3)
        self.assertEqual(result["retained_row_mismatches"], 0)

    def test_duplicate_replacing_lower_value_keeps_count_max_but_fails(self):
        self.standard([(5, 0, "北"), (4, 0, "北"), (4, 0, "北")])
        result = self.inspect()
        self.assertEqual(result["direction_mapping_mismatches"], 0)
        self.assertEqual(result["detail_max_summary_mismatches"], 0)
        self.assertEqual(result["detail_rows"], 3)
        self.assertEqual(result["retained_row_mismatches"], 1)

    def test_missing_extra_and_wrong_order_rows_fail(self):
        for rows in ([(5, 0, "北"), (4, 0, "北")],
                     [(5, 0, "北"), (4, 0, "北"), (3, 0, "北"), (3, 0, "北")],
                     [(5, 0, "北"), (3, 0, "北"), (4, 0, "北")]):
            with self.subTest(rows=rows):
                self.standard(rows)
                self.assertGreater(self.inspect()["retained_row_mismatches"], 0)

    def test_sub_nanometer_change_does_not_hide_in_old_broad_tolerance(self):
        self.standard([(5, 0, "北"), (4 + 1e-10, 0, "北"), (3, 0, "北")])
        self.assertEqual(self.inspect()["retained_row_mismatches"], 1)

    def test_only_actual_numeric_xml_serialization_is_normalized(self):
        value = 1.2345678901234567
        self.source([(value, 0, 0, 1)])
        self.workbook([(value, 0, "北")], maximum=value)
        self.assertEqual(self.inspect()["retained_row_mismatches"], 0)

    def test_period_message_cap_and_blank_exclusion_rules(self):
        self.source([(5, 0, 0, 1), (4, 0, 43199, 4), (3, 0, 43200, 1),
                     (2, 0, 0, 2), (6, 0, 0, 1), (499, 123.75, 0, 1), (501, 123.75, 0, 1)])
        self.workbook([(5, 0, "北"), (4, 0, "北"), (499, 123.75, "東南微東")])
        result = self.inspect()
        self.assertEqual(result["retained_row_mismatches"], 0)
        self.assertEqual(result["retained_expected_rows"], 3)

    def test_spool_truncation_never_becomes_empty_pass(self):
        self.standard()
        with (self.spool / "20260601.bin").open("ab") as handle:
            handle.write(b"unexpected")
        with self.assertRaisesRegex(ValueError, "length/count"):
            self.inspect()

    def test_production_direction_mutation_does_not_redefine_oracle(self):
        self.standard()
        with mock.patch("ais_monthly_app.DIRECTION_ORDER", ["wrong"] * 32), \
                mock.patch("ais_monthly_app.degree_to_direction_index", return_value=(1, 0)):
            self.assertEqual(self.inspect()["direction_mapping_mismatches"], 0)

    def test_bool_is_not_a_numeric_zero(self):
        self.source([(5, 0, 0, 1), (0, 0, 0, 1)])
        self.workbook([(5, 0, "北"), (False, 0, "北")])
        self.assertEqual(self.inspect()["retained_row_mismatches"], 1)


if __name__ == "__main__":
    unittest.main()
