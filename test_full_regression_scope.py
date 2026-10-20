from __future__ import annotations

import contextlib
import copy
import io
import json
import os
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, datetime
from pathlib import Path
from unittest import mock

from scripts import run_full_regression as runner
from scripts import validate_full_regression as validator


class FullRegressionScopeTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory(prefix="ais-regression-scope-")
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.source = self.root / "source"
        self.output = self.root / "output"
        for token, port, month in (("HWLN_202604", "HWLN", 4), ("KLNG_202605", "KLNG", 5)):
            folder = self.source / token
            folder.mkdir(parents=True)
            (folder / f"D&TMOK {port}_2026{month:02d}01_any.xlsx").write_bytes(b"scope-only")
            (folder / f"D&TMOK {port}_2026{month:02d}02_any.xlsx").write_bytes(b"scope-only")
        self.plan = [runner.serializable_job(job) for job in runner.discover_jobs(self.source)]
        self.records = [{**copy.deepcopy(job), "current_parser_logical_days": job["tolerant_logical_days"],
                         "status": "PASS", "modern": str(self.output / job["token"] / "modern.xlsx"),
                         "delivery": str(self.output / job["token"] / "delivery")}
                        for job in self.plan]
        self.report = {"source_root": str(self.source), "scope": "FULL", "status": "COMPLETE",
                       "planned_jobs": self.plan, "expected_job_count": len(self.plan),
                       "jobs": self.records}

    def audit(self, report: dict | None = None, source_root: Path | None = None) -> dict:
        return validator.audit_run_scope(self.report if report is None else report, source_root)

    def write_report(self, value: dict) -> None:
        self.output.mkdir(exist_ok=True)
        (self.output / "run_report.json").write_text(json.dumps(value), encoding="utf-8")

    def invoke_runner(self, *extra: str) -> int:
        with contextlib.redirect_stdout(io.StringIO()):
            return runner.main([str(self.source), str(self.output), "--workers", "1", *extra])

    def test_complete_plan_and_records_match(self) -> None:
        result = self.audit()
        self.assertTrue(result["complete"], result["issues"])
        self.assertEqual(result["expected_job_count"], 2)

    def test_empty_legacy_report_cannot_pass(self) -> None:
        self.write_report({"jobs": []})
        result = validator.validate_run(self.output)
        self.assertFalse(result["pass"])
        self.assertEqual(result["workbook_count"], 0)

    def test_nonempty_legacy_without_plan_or_explicit_source_root_cannot_pass(self) -> None:
        legacy = {"jobs": copy.deepcopy(self.records)}
        self.assertFalse(self.audit(legacy)["complete"])
        self.assertTrue(self.audit(legacy, self.source)["complete"])

    def test_legacy_one_completed_job_of_two_is_not_full_pass(self) -> None:
        legacy = {"jobs": [self.records[0]]}
        result = self.audit(legacy, self.source)
        self.assertFalse(result["complete"])
        self.assertEqual(result["missing_jobs"], [self.records[1]["token"]])

    def test_missing_planned_job_cannot_pass(self) -> None:
        report = copy.deepcopy(self.report)
        report["jobs"].pop()
        self.assertFalse(self.audit(report)["complete"])

    def test_duplicate_recorded_job_cannot_substitute_for_missing_job(self) -> None:
        report = copy.deepcopy(self.report)
        report["jobs"] = [report["jobs"][0], report["jobs"][0]]
        result = self.audit(report)
        self.assertFalse(result["complete"])
        self.assertTrue(any("duplicate recorded" in issue for issue in result["issues"]))

    def test_duplicate_planned_tokens_are_invalid(self) -> None:
        report = copy.deepcopy(self.report)
        report["planned_jobs"] = [self.plan[0], self.plan[0]]
        result = self.audit(report)
        self.assertFalse(result["complete"])
        self.assertTrue(any("duplicate planned" in issue for issue in result["issues"]))

    def test_in_progress_interrupted_and_failed_reports_are_not_complete(self) -> None:
        for status in ("IN_PROGRESS", "INTERRUPTED", "FAILED"):
            with self.subTest(status=status):
                report = {**self.report, "status": status}
                self.assertFalse(self.audit(report)["complete"])

    def test_subset_is_not_full_regression(self) -> None:
        report = {**self.report, "scope": "SUBSET"}
        self.assertFalse(self.audit(report)["complete"])

    def test_wrong_expected_job_count_cannot_pass(self) -> None:
        report = {**self.report, "expected_job_count": 1}
        self.assertFalse(self.audit(report)["complete"])

    def test_wrong_physical_day_and_parser_counts_cannot_pass(self) -> None:
        for field in ("physical_files", "tolerant_logical_days", "current_parser_logical_days"):
            with self.subTest(field=field):
                report = copy.deepcopy(self.report)
                report["jobs"][0][field] -= 1
                self.assertFalse(self.audit(report)["complete"])

    def test_same_count_but_wrong_actual_day_names_cannot_pass(self) -> None:
        report = copy.deepcopy(self.report)
        report["jobs"][0]["logical_days"][1] = "2026-04-03"
        self.assertFalse(self.audit(report)["complete"])

    def test_same_count_but_wrong_source_file_names_cannot_pass(self) -> None:
        report = copy.deepcopy(self.report)
        report["jobs"][0]["source_files"][1] = "D&TMOK HWLN_20260403_any.xlsx"
        self.assertFalse(self.audit(report)["complete"])

    def test_saved_plan_must_match_explicit_current_source_scope(self) -> None:
        report = copy.deepcopy(self.report)
        report["planned_jobs"].pop()
        report["jobs"].pop()
        report["expected_job_count"] = 1
        self.assertFalse(self.audit(report, self.source)["complete"])

    def test_unavailable_source_root_cannot_be_used_as_scope_evidence(self) -> None:
        result = self.audit({"jobs": self.records}, self.root / "missing")
        self.assertFalse(result["complete"])

    def test_runner_saves_full_plan_before_first_job_and_checkpoints_each_job(self) -> None:
        seen = []
        def job(job: dict, output: Path, workers: int) -> dict:
            progress = json.loads((self.output / "run_report.json").read_text(encoding="utf-8"))
            self.assertEqual(progress["status"], "IN_PROGRESS")
            self.assertEqual(progress["expected_job_count"], 2)
            self.assertEqual(len(progress["planned_jobs"]), 2)
            self.assertEqual(len(progress["jobs"]), len(seen))
            seen.append(runner.job_token(job))
            return {**runner.serializable_job(job), "status": "PASS",
                    "current_parser_logical_days": job["tolerant_logical_days"]}
        with mock.patch.object(runner, "run_job", side_effect=job):
            self.assertEqual(self.invoke_runner(), 0)
        final = json.loads((self.output / "run_report.json").read_text(encoding="utf-8"))
        self.assertEqual(final["status"], "COMPLETE")
        self.assertTrue(self.audit(final)["complete"])

    def test_runner_interruption_preserves_plan_and_incomplete_status(self) -> None:
        calls = 0
        def job(job: dict, output: Path, workers: int) -> dict:
            nonlocal calls
            calls += 1
            if calls == 2:
                raise KeyboardInterrupt("synthetic interruption")
            return {**runner.serializable_job(job), "status": "PASS",
                    "current_parser_logical_days": job["tolerant_logical_days"]}
        with mock.patch.object(runner, "run_job", side_effect=job):
            with self.assertRaises(KeyboardInterrupt):
                self.invoke_runner()
        interrupted = json.loads((self.output / "run_report.json").read_text(encoding="utf-8"))
        self.assertEqual(interrupted["status"], "INTERRUPTED")
        self.assertEqual(len(interrupted["jobs"]), 1)
        self.assertEqual(interrupted["expected_job_count"], 2)
        self.assertFalse(validator.validate_run(self.output)["pass"])

    def test_worker_cannot_mutate_saved_source_plan_by_shared_list(self) -> None:
        initial_days = list(self.plan[0]["logical_days"])
        def job(job: dict, output: Path, workers: int) -> dict:
            job["logical_days"][0] = "2026-04-09"
            return {**runner.serializable_job(job), "status": "PASS",
                    "current_parser_logical_days": job["tolerant_logical_days"]}
        with mock.patch.object(runner, "run_job", side_effect=job):
            self.assertEqual(self.invoke_runner(), 0)
        final = json.loads((self.output / "run_report.json").read_text(encoding="utf-8"))
        self.assertEqual(final["planned_jobs"][0]["logical_days"], initial_days)
        self.assertFalse(self.audit(final)["complete"])

    def test_new_attempt_cannot_reuse_stale_complete_report(self) -> None:
        self.write_report(self.report)
        with mock.patch.object(runner, "run_job", side_effect=KeyboardInterrupt("first job")):
            with self.assertRaises(KeyboardInterrupt):
                self.invoke_runner()
        interrupted = json.loads((self.output / "run_report.json").read_text(encoding="utf-8"))
        self.assertEqual(interrupted["jobs"], [])
        self.assertEqual(interrupted["status"], "INTERRUPTED")
        self.assertFalse(self.audit(interrupted)["complete"])

    def test_empty_source_is_failed_and_never_returns_success(self) -> None:
        for path in self.source.rglob("*.xlsx"):
            path.unlink()
        with self.assertRaises(ValueError):
            self.invoke_runner()
        report = json.loads((self.output / "run_report.json").read_text(encoding="utf-8"))
        self.assertEqual(report["status"], "FAILED")
        self.assertFalse(self.audit(report)["complete"])

    def test_duplicate_source_month_folders_are_not_silently_overwritten(self) -> None:
        duplicate = self.source / "another_HWln_April"
        duplicate.mkdir()
        (duplicate / "D&TMOK HWLN_20260403_any.xlsx").write_bytes(b"scope-only")
        with self.assertRaisesRegex(ValueError, "相同港別"):
            runner.discover_jobs(self.source)

    def test_malformed_ais_source_name_cannot_disappear_from_plan(self) -> None:
        (self.source / "D&TMOK HWLN_BADDATE.xlsx").write_bytes(b"scope-only")
        with self.assertRaisesRegex(ValueError, "不能從全量計畫"):
            runner.discover_jobs(self.source)

    def test_runner_cannot_create_artifacts_in_or_above_source_tree(self) -> None:
        before = sorted(str(path.relative_to(self.source)) for path in self.source.rglob("*"))
        for target in (self.source, self.source / "new-output", self.root):
            with self.subTest(target=target):
                with self.assertRaisesRegex(ValueError, "完全分離"):
                    runner.main([str(self.source), str(target)])
        after = sorted(str(path.relative_to(self.source)) for path in self.source.rglob("*"))
        self.assertEqual(before, after)
        self.assertFalse((self.source / "run_report.json").exists())

    def test_runner_report_hard_link_to_raw_source_is_rejected(self) -> None:
        raw = next(self.source.rglob("*.xlsx"))
        before = raw.read_bytes()
        self.output.mkdir()
        os.link(raw, self.output / "run_report.json")
        with self.assertRaisesRegex(ValueError, "hard-link"):
            self.invoke_runner()
        self.assertEqual(raw.read_bytes(), before)

    def test_validator_outputs_cannot_modify_raw_run_or_reference_trees(self) -> None:
        self.write_report(self.report)
        reference = self.root / "reference"
        reference.mkdir()
        raw = next(self.source.rglob("*.xlsx"))
        before_raw = raw.read_bytes()
        before_report = (self.output / "run_report.json").read_bytes()
        for target in (raw, self.source / "audit.json", self.output / "audit.json",
                       self.output / "run_report.json", reference / "audit.json"):
            with self.subTest(target=target):
                with self.assertRaises(ValueError):
                    validator.guard_validation_output(target, self.output, reference, self.source)
        self.assertEqual(raw.read_bytes(), before_raw)
        self.assertEqual((self.output / "run_report.json").read_bytes(), before_report)
        self.assertFalse((reference / "audit.json").exists())

    def test_validator_json_hard_link_to_workbook_is_rejected(self) -> None:
        self.write_report(self.report)
        workbook = self.output / "workbook.xlsx"
        workbook.write_bytes(b"workbook-protected")
        alias = self.root / "alias.json"
        os.link(workbook, alias)
        with self.assertRaisesRegex(ValueError, "hard-link"):
            validator.guard_validation_output(alias, self.output, None, self.source)
        self.assertEqual(workbook.read_bytes(), b"workbook-protected")

    def test_validator_safe_separate_json_output_is_allowed(self) -> None:
        self.write_report(self.report)
        validator.guard_validation_output(self.root / "audit.json", self.output, None, self.source)
        self.assertFalse((self.root / "audit.json").exists())


class ReferenceDeliveryLayoutTests(unittest.TestCase):
    # Portable synthetic XML fixtures independently copied from the accepted
    # reference geometry, never fetched from producer/validator layout tables.
    A = (((12, 9524, 1, 19050), (22, 609599, 30, 190500)),
         ((0, 9525, 43, 0), (14, 590550, 56, 142875)))
    B = (((11, 838200, 1, 33336), (23, 0, 31, 9525)),
         ((0, 0, 43, 14287), (15, 9524, 56, 157162)))
    I = (((12, 0, 1, 9524), (23, 9525, 31, 19049)),
         ((0, 0, 43, 9525), (14, 590549, 56, 152400)))
    LEGENDS = {
        "a": (0.76830972615675164, 0.10649836959552665, 0.22035882908404159, 0.83653926912995602),
        "b": (0.8682679434523134, 0.11656551957448454, 0.12020467902607275, 0.82470676043372704),
        "integrated": (0.87979532097392721, 0.12502090170582877, 0.10867730150445892, 0.82013252305426954)}
    NAMES = {"full_a": "KLNG 6月 32方位數值_2.xlsx", "full_b": "KLNG 6月 32方位數值.xlsx",
             "summary_a": "KLNG 6月 32方位數值_2 總表.xlsx", "summary_b": "KLNG 6月 32方位數值 總表.xlsx",
             "integrated": "KLNG_6月_32方位_每日較大值整合總表.xlsx"}

    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix="ais-reference-layout-")
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)

    def fixture(self, role="summary_b", days=30):
        def element(parent, tag, **attributes):
            prefix, name = tag.split(":")
            return ET.SubElement(parent, f"{{{validator.NS[prefix]}}}{name}", attributes)
        def manual(parent, values):
            layout = element(parent, "c:layout")
            if values is not None:
                node = element(layout, "c:manualLayout")
                for tag, value in (("xMode", "edge"), ("yMode", "edge"), *zip(("x", "y", "w", "h"), values)):
                    element(node, "c:" + tag, val=str(value))
        scope = "integrated" if role == "integrated" else role[-1]
        anchors = self.I if scope == "integrated" else self.A if scope == "a" else self.B
        drawing = ET.Element(f"{{{validator.XDR_NS}}}wsDr")
        for index, vertices in enumerate(anchors):
            anchor = element(drawing, "xdr:twoCellAnchor")
            for edge, values in zip(("from", "to"), vertices):
                vertex = element(anchor, "xdr:" + edge)
                for pos, (tag, value) in enumerate(zip(("col", "colOff", "row", "rowOff"), values)):
                    element(vertex, "xdr:" + tag).text = str(value + (days - 30 if index == 1 and pos == 2 else 0))
            element(anchor, "c:chart", **{f"{{{validator.R_NS}}}id": f"rId{index + 1}"})
        sheet = ET.Element(f"{{{validator.S_NS}}}worksheet")
        formatting = element(sheet, "s:sheetFormatPr", defaultRowHeight="15.75")
        if scope != "a":
            formatting.set("defaultColWidth", "9")
        columns = element(sheet, "s:cols")
        widths = [] if scope == "a" else [(1, 1, "14.5703125")]
        if role == "full_b":
            widths += [(2, 6, "12.7109375"), (7, 7, "11.5703125"), (8, 12, "12.7109375"), (24, 33, "12.7109375")]
        for first, last, width in widths:
            element(columns, "s:col", min=str(first), max=str(last), width=width)
        element(columns, "s:col", min="35", max="38", width="0", hidden="1")
        element(sheet, "s:sheetData")
        charts = []
        for index in range(2):
            root = ET.Element(f"{{{validator.C_NS}}}chartSpace")
            element(root, "c:lang", val="zh-TW")
            parent_text = element(root, "c:txPr")
            paragraph = element(parent_text, "a:p")
            element(element(paragraph, "a:pPr"), "a:defRPr")
            element(paragraph, "a:endParaRPr", lang="zh-TW")
            chart = element(root, "c:chart")
            title = element(chart, "c:title")
            properties = element(title, "a:defRPr", sz="1400")
            for script, face in (("latin", "+mn-lt"), ("ea", "+mn-ea"), ("cs", "+mn-cs")):
                element(properties, f"a:{script}", typeface="Calibri" if scope == "integrated" else face)
            element(element(properties, "a:solidFill"), "a:srgbClr", val="595959")
            if index == 1 and scope == "a":
                title_manual = element(element(title, "c:layout"), "c:manualLayout")
                for tag, value in (("xMode", "edge"), ("yMode", "edge"),
                                   ("x", "0.3523162112572919"), ("y", "0.027777777777777776")):
                    element(title_manual, f"c:{tag}", val=value)
            plot = element(chart, "c:plotArea")
            manual(plot, (0.064771103328797777, 0.17572538384888339, 0.65954183489103524, 0.74501858184039171)
                   if index == 0 and scope == "a" else None)
            if index == 0 and scope == "a":
                element(plot.find("c:layout/c:manualLayout", validator.NS), "c:layoutTarget", val="inner")
            element(plot, "c:radarChart" if index == 0 else "c:lineChart")
            if index == 1:
                columns = element(plot, "c:barChart")
                for tag, value in (("grouping", "clustered"), ("gapWidth", "219"), ("overlap", "-27")):
                    element(columns, f"c:{tag}", val=value)
                for axis_tag in ("catAx", "valAx"):
                    axis = element(plot, f"c:{axis_tag}")
                    for mark in ("majorTickMark", "minorTickMark"):
                        element(axis, f"c:{mark}", val="none")
                    if axis_tag == "catAx":
                        baseline = element(element(axis, "c:spPr"), "a:ln", w="9525")
                        element(element(baseline, "a:solidFill"), "a:srgbClr", val="D9D9D9")
            legend = element(chart, "c:legend")
            element(legend, "c:legendPos", val="r" if index == 0 else "b")
            manual(legend, self.LEGENDS[scope] if index == 0 else None)
            properties = element(legend, "a:defRPr", sz="900")
            for script, face in (("latin", "+mn-lt"), ("ea", "+mn-ea"), ("cs", "+mn-cs")):
                element(properties, f"a:{script}", typeface="Calibri" if scope == "integrated" else face)
            element(element(properties, "a:solidFill"), "a:srgbClr", val="595959")
            charts.append(root)
        return drawing, sheet, charts

    def audit(self, fixture, role="summary_b", days=30, swapped=False, theme_faces=None):
        drawing, sheet, charts = fixture
        path = self.root / self.NAMES[role]
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("xl/drawings/drawing1.xml", ET.tostring(drawing))
            archive.writestr("xl/worksheets/sheet31.xml", ET.tostring(sheet))
            archive.writestr("xl/worksheets/_rels/sheet31.xml.rels", '<Relationships><Relationship Id="rId1" Type="x/drawing" Target="../drawings/drawing1.xml"/></Relationships>')
            targets = (2, 1) if swapped else (1, 2)
            archive.writestr("xl/drawings/_rels/drawing1.xml.rels", '<Relationships>' + ''.join(
                f'<Relationship Id="rId{i}" Type="x/chart" Target="../charts/chart{target}.xml"/>'
                for i, target in zip((1, 2), targets)) + '</Relationships>')
            for index, chart in enumerate(charts, 1):
                archive.writestr(f"xl/charts/chart{index}.xml", ET.tostring(chart))
            minor_faces = theme_faces or (("Calibri", "Calibri", "Calibri")
                                          if role == "integrated" else ("Calibri", "", ""))
            theme = ET.Element(f"{{{validator.A_NS}}}theme")
            elements = ET.SubElement(theme, f"{{{validator.A_NS}}}themeElements")
            scheme = ET.SubElement(elements, f"{{{validator.A_NS}}}fontScheme")
            minor = ET.SubElement(scheme, f"{{{validator.A_NS}}}minorFont")
            for script, face in zip(("latin", "ea", "cs"), minor_faces):
                ET.SubElement(minor, f"{{{validator.A_NS}}}{script}", typeface=face)
            if role != "integrated":
                ET.SubElement(minor, f"{{{validator.A_NS}}}font", script="Hant", typeface="新細明體")
            archive.writestr("xl/theme/theme1.xml", ET.tostring(theme))
        with zipfile.ZipFile(path) as archive:
            return validator.validate_reference_layout(archive, path, days, drawing, *charts)

    def test_all_five_reference_roles_and_month_lengths(self):
        for role in self.NAMES:
            for days in (28, 30, 31):
                with self.subTest(role=role, days=days):
                    self.assertEqual(self.audit(self.fixture(role, days), role, days), [])

    def test_full_b_cannot_use_summary_b_column_geometry(self):
        issues = self.audit(self.fixture("summary_b"), "full_b")
        self.assertTrue(any("column geometry" in issue for issue in issues))

    def test_summary_b_cannot_be_mislabelled_full_b(self):
        self.assertTrue(self.audit(self.fixture("full_b"), "summary_b"))

    def test_one_emu_to_vertex_error_is_detected(self):
        fixture = self.fixture()
        fixture[0].find("xdr:twoCellAnchor/xdr:to/xdr:colOff", validator.NS).text = "1"
        self.assertTrue(any("to anchor" in issue for issue in self.audit(fixture)))

    def test_statistics_to_vertex_must_also_shift_with_actual_days(self):
        fixture = self.fixture(days=28)
        fixture[0].findall("xdr:twoCellAnchor", validator.NS)[1].find("xdr:to/xdr:row", validator.NS).text = "56"
        self.assertTrue(any("chart 2 to anchor" in issue for issue in self.audit(fixture, days=28)))

    def test_reference_default_row_and_column_widths_are_required(self):
        fixture = self.fixture()
        fixture[1].find("s:sheetFormatPr", validator.NS).set("defaultRowHeight", "15")
        fixture[1].find("s:sheetFormatPr", validator.NS).set("defaultColWidth", "8.43")
        issues = self.audit(fixture)
        self.assertTrue(any("row height" in issue for issue in issues))
        self.assertTrue(any("default column" in issue for issue in issues))

    def test_title_legend_sizes_and_text_color_are_checked(self):
        fixture = self.fixture()
        fixture[2][0].find(".//c:title/a:defRPr", validator.NS).set("sz", "1600")
        fixture[2][0].find(".//c:legend/a:defRPr", validator.NS).set("sz", "1000")
        fixture[2][1].find(".//c:title//a:srgbClr", validator.NS).set("val", "000000")
        issues = self.audit(fixture)
        self.assertTrue(any("14pt" in issue for issue in issues))
        self.assertTrue(any("9pt" in issue for issue in issues))
        self.assertTrue(any("text color" in issue for issue in issues))

    def test_wrong_radar_manual_legend_is_not_accepted(self):
        fixture = self.fixture()
        fixture[2][0].find(".//c:legend/c:layout/c:manualLayout/c:x", validator.NS).set("val", "0.7")
        self.assertTrue(any("legend manual layout" in issue for issue in self.audit(fixture)))

    def test_radar_plot_inner_target_cannot_silently_become_outer(self):
        fixture = self.fixture("full_a")
        fixture[2][0].find(".//c:plotArea/c:layout/c:manualLayout/c:layoutTarget", validator.NS).set("val", "outer")
        self.assertTrue(any("target must be inner" in issue for issue in self.audit(fixture, "full_a")))

    def test_swapped_chart_relationships_are_rejected_even_with_correct_anchors(self):
        self.assertTrue(any("wrong chart role" in issue for issue in self.audit(self.fixture(), swapped=True)))

    def test_integrated_february_name_is_not_period_a(self):
        self.assertEqual(validator.delivery_artifact_role(Path("HWLN_2月_32方位_每日較大值整合總表.xlsx")), "integrated")

    def test_statistics_gap_and_overlap_defaults_are_rejected(self):
        fixture = self.fixture()
        columns = fixture[2][1].find(".//c:barChart", validator.NS)
        columns.find("c:gapWidth", validator.NS).set("val", "150")
        columns.find("c:overlap", validator.NS).set("val", "0")
        self.assertTrue(any("gap/overlap" in issue for issue in self.audit(fixture)))

    def test_integrated_theme_cannot_silently_use_ab_script_faces(self):
        self.assertTrue(any("script font faces" in issue for issue in self.audit(
            self.fixture("integrated"), "integrated", theme_faces=("Calibri", "", ""))))

    def test_literal_font_and_us_parent_language_are_rejected(self):
        fixture = self.fixture()
        fixture[2][0].find(".//c:title/a:defRPr/a:latin", validator.NS).set("typeface", "新細明體")
        fixture[2][0].find("c:lang", validator.NS).set("val", "en-US")
        issues = self.audit(fixture)
        self.assertTrue(any("font tokens" in issue for issue in issues))
        self.assertTrue(any("chart language" in issue for issue in issues))

    def test_integrated_theme_tokens_without_native_compatibility_are_rejected(self):
        fixture = self.fixture("integrated")
        properties = fixture[2][0].find(".//c:title/a:defRPr", validator.NS)
        for script, face in (("latin", "+mn-lt"), ("ea", "+mn-ea"), ("cs", "+mn-cs")):
            properties.find(f"a:{script}", validator.NS).set("typeface", face)
        self.assertTrue(any("font tokens" in issue for issue in self.audit(fixture, "integrated")))

    def test_statistics_axis_ticks_baseline_and_manual_title_are_checked(self):
        fixture = self.fixture("full_a")
        fixture[2][1].find(".//c:catAx/c:majorTickMark", validator.NS).set("val", "out")
        fixture[2][1].find(".//c:catAx/c:spPr/a:ln/a:solidFill/a:srgbClr", validator.NS).set("val", "BFBFBF")
        fixture[2][1].find("c:chart/c:title/c:layout/c:manualLayout/c:x", validator.NS).set("val", "0.5")
        issues = self.audit(fixture, "full_a")
        self.assertTrue(any("axis tick marks" in issue for issue in issues))
        self.assertTrue(any("baseline" in issue for issue in issues))
        self.assertTrue(any("title position" in issue for issue in issues))


class SemanticPackageProvenanceTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix="ais-provenance-")
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)

    def write_modern(self, folder, *, number=12.5, formula="=SUM(A1:A2)", color="#FF0000",
                     self_path=None, guide_label="Modern analysis", creator="AIS QA",
                     source="C:\\protected\\source", unexpected_consumer=False, row_height=30,
                     timings=(1.25, 2.5), timing_formula=None, timing_header="秒數", count=100,
                     log_date_offset=0):
        import xlsxwriter
        path = self.root / folder / "modern.xlsx"
        path.parent.mkdir(parents=True, exist_ok=True)
        workbook = xlsxwriter.Workbook(path)
        workbook.set_properties({"title": "AIS numeric proof", "author": creator,
                                 "created": datetime(2026, 6, 1 if folder.endswith("a") else 2)})
        guide = workbook.add_worksheet("操作說明")
        value = str(path.resolve()) if self_path is None else self_path
        guide.write("A7", guide_label)
        guide.write("B7", value)
        guide.set_row(6, row_height)
        system = workbook.add_worksheet("系統資料")
        system.write("A11", "Analysis workbook path")
        system.write("B11", value)
        system.write("A12", "Input directory")
        system.write("B12", source)
        total = workbook.add_worksheet("總表")
        total.write("A4", "日期")
        log = workbook.add_worksheet("處理紀錄")
        log.write_row("A5", ["日期", "來源檔", "掃描列數", "合格列數", "錯誤列數",
                              "訊息類型略過", "非 East 略過", timing_header])
        date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})
        for offset, timing in enumerate(timings):
            day = datetime(2026, 6, 1 + offset)
            total.write_datetime(4 + offset, 0, day, date_format)
            log.write_datetime(5 + offset, 0, datetime(2026, 6, 1 + offset + log_date_offset), date_format)
            log.write(5 + offset, 1, "D&TMOK source.xlsx")
            log.write_row(5 + offset, 2, [count, 90, 0, 10, 0])
            if timing_formula is not None:
                log.write_formula(5 + offset, 7, timing_formula)
            else:
                log.write(5 + offset, 7, timing)
        data = workbook.add_worksheet("Data")
        data.write("A1", number)
        data.write("A2", 3.0)
        data.write_formula("B1", formula)
        if unexpected_consumer:
            data.write("C1", value)
        chart = workbook.add_chart({"type": "line"})
        chart.add_series({"values": "=Data!$A$1:$A$2", "line": {"color": color}})
        data.insert_chart("D1", chart)
        workbook.close()
        return path

    def rewrite_core(self, path, field, value):
        with zipfile.ZipFile(path) as archive:
            parts = {item.filename: archive.read(item.filename) for item in archive.infolist()}
        core = ET.fromstring(parts["docProps/core.xml"])
        core.find("{http://purl.org/dc/terms/}" + field).text = value
        parts["docProps/core.xml"] = ET.tostring(core)
        with zipfile.ZipFile(path, "w") as archive:
            for name, payload in parts.items():
                archive.writestr(name, payload)

    def test_only_verified_self_paths_and_real_creation_timestamps_are_normalized(self):
        first = self.write_modern("fresh-a")
        repeat = self.write_modern("repeat-b-longer-root", timings=(2.345, 3.456))
        left = validator.semantic_package_fingerprint(first)
        right = validator.semantic_package_fingerprint(repeat)
        self.assertEqual(left["sha256"], right["sha256"])
        self.assertEqual(len(left["provenance_normalizations"]), 6)
        self.assertEqual({item["field"] for item in left["provenance_normalizations"]},
                         {"created", "modified", "操作說明!B7", "系統資料!B11", "處理紀錄!H6", "處理紀錄!H7"})
        self.assertNotEqual(left["provenance_normalizations"], right["provenance_normalizations"])

    def test_same_length_numeric_change_is_not_normalized(self):
        self.assertNotEqual(validator.semantic_package_hash(self.write_modern("first-a")),
                            validator.semantic_package_hash(self.write_modern("next-b", number=13.5)))

    def test_same_length_formula_change_is_not_normalized(self):
        self.assertNotEqual(validator.semantic_package_hash(self.write_modern("first-a")),
                            validator.semantic_package_hash(self.write_modern("next-b", formula="=MAX(A1:A2)")))

    def test_same_length_chart_color_change_is_not_normalized(self):
        self.assertNotEqual(validator.semantic_package_hash(self.write_modern("first-a")),
                            validator.semantic_package_hash(self.write_modern("next-b", color="#00FF00")))

    def test_source_root_difference_is_not_normalized(self):
        self.assertNotEqual(validator.semantic_package_hash(self.write_modern("first-a")),
                            validator.semantic_package_hash(self.write_modern("next-b", source="D:\\protected\\source")))

    def test_core_creator_change_is_not_discarded_with_timestamps(self):
        self.assertNotEqual(validator.semantic_package_hash(self.write_modern("first-a")),
                            validator.semantic_package_hash(self.write_modern("next-b", creator="AIS QB")))

    def test_wrong_same_length_self_path_fails_closed(self):
        path = self.root / "fresh-a" / "modern.xlsx"
        other = str((self.root / "rogue-a" / "modern.xlsx").resolve())
        self.assertEqual(len(str(path.resolve())), len(other))
        self.write_modern("fresh-a", self_path=other)
        with self.assertRaisesRegex(ValueError, "unexpected self workbook path"):
            validator.semantic_package_hash(path)

    def test_wrong_provenance_label_fails_closed(self):
        path = self.write_modern("fresh-a", guide_label="Modern analyses")
        with self.assertRaisesRegex(ValueError, "wrong provenance label"):
            validator.semantic_package_hash(path)

    def test_shared_self_path_used_in_another_cell_fails_closed(self):
        path = self.write_modern("fresh-a", unexpected_consumer=True)
        with self.assertRaisesRegex(ValueError, "unexpected consumers"):
            validator.semantic_package_hash(path)

    def test_row_geometry_is_not_discarded_even_in_provenance_row(self):
        self.assertNotEqual(validator.semantic_package_hash(self.write_modern("first-a")),
                            validator.semantic_package_hash(self.write_modern("next-b", row_height=22)))

    def test_malformed_timestamp_does_not_become_ignored_metadata(self):
        path = self.write_modern("fresh-a")
        self.rewrite_core(path, "created", "2026-99-99T25:00:00Z")
        with self.assertRaises(ValueError):
            validator.semantic_package_hash(path)

    def test_timing_formula_is_not_ignored_as_performance_provenance(self):
        path = self.write_modern("fresh-a", timing_formula="=SUM(A1:A2)")
        with self.assertRaisesRegex(ValueError, "plain numeric provenance"):
            validator.semantic_package_hash(path)

    def test_negative_or_nonfinite_elapsed_values_fail_closed(self):
        path = self.write_modern("fresh-a", timings=(-1.25, 2.5))
        with self.assertRaisesRegex(ValueError, "invalid elapsed seconds"):
            validator.semantic_package_hash(path)

    def test_wrong_timing_header_cannot_hide_numeric_data(self):
        path = self.write_modern("fresh-a", timing_header="DISTANCE")
        with self.assertRaisesRegex(ValueError, "wrong timing log header"):
            validator.semantic_package_hash(path)

    def test_elapsed_row_must_match_real_logical_day(self):
        path = self.write_modern("fresh-a", log_date_offset=1)
        with self.assertRaisesRegex(ValueError, "timing logical day mapping mismatch"):
            validator.semantic_package_hash(path)

    def test_source_row_counts_are_not_ignored_with_elapsed_seconds(self):
        self.assertNotEqual(validator.semantic_package_hash(self.write_modern("first-a")),
                            validator.semantic_package_hash(self.write_modern("next-b", count=101)))


class ReferenceRadarOracleTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix="ais-radar-oracle-")
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)

    def write_sparse_summary(self, calendar_days=(1, 2, 3, 4), observed_indexes=(0, 2)):
        import ais_monthly_app as app
        config = app.AppConfig(self.root, self.root / "modern.xlsx", "HWLN", 2026, 6)
        days = [app.DayResult(date(2026, 6, day), (), {}) for day in calendar_days]
        values = {(days[index].day, "北"): 100.25 + index for index in observed_indexes}
        path = self.root / "HWLN 6月 32方位數值_2 總表.xlsx"
        app._write_summary_workbook(path, config, days, values, scope_title="Period A",
                                    artifact_role="summary_a")
        return path

    def test_skipped_empty_day_retains_logical_day_palette_not_series_position(self):
        path = self.write_sparse_summary()
        with zipfile.ZipFile(path) as archive:
            chart = ET.fromstring(archive.read("xl/charts/chart1.xml"))
        series = chart.findall(".//c:radarChart/c:ser", validator.NS)
        self.assertEqual([item.find("c:tx/c:v", validator.NS).text for item in series],
                         ["6月1日", "6月3日"])
        self.assertEqual([item.find("c:spPr/a:solidFill/a:srgbClr", validator.NS).get("val")
                          for item in series], ["5B9BD5", "A5A5A5"])
        self.assertEqual(validator.validate_delivery_charts(path, 4)["issues"], [])

    def test_palette_uses_ordered_source_day_plan_not_calendar_day_number(self):
        path = self.write_sparse_summary(calendar_days=(8, 10), observed_indexes=(0, 1))
        self.assertEqual(validator.validate_delivery_charts(path, 2)["issues"], [])

    def test_producer_palette_mutation_does_not_redefine_reference_oracle(self):
        import ais_monthly_app as app
        with mock.patch.object(app, "LEGACY_RADAR_PALETTE", ("#123456",) * 30):
            path = self.write_sparse_summary()
        self.assertTrue(any(" color" in issue for issue in
                            validator.validate_delivery_charts(path, 4)["issues"]))
        self.assertEqual(validator.REFERENCE_RADAR_PALETTE[:3], ("5B9BD5", "ED7D31", "A5A5A5"))

    def test_producer_abbreviation_mutation_does_not_redefine_reference_oracle(self):
        import ais_monthly_app as app
        with mock.patch.object(app, "LEGACY_DIRECTION_ABBREVIATIONS", ["WRONG"] * 32):
            path = self.write_sparse_summary()
        self.assertTrue(any("category order" in issue for issue in
                            validator.validate_delivery_charts(path, 4)["issues"]))
        with self.assertRaisesRegex(AssertionError, "direction headers"):
            validator.read_summary(path, 4)


if __name__ == "__main__":
    unittest.main()
