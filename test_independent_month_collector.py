import contextlib
import io
import json
import os
from pathlib import Path
import tempfile
import unittest

from scripts import independent_source_audit as independent
from scripts import run_independent_month_audits as collector


class IndependentMonthCollectorTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix="ais-month-collector-")
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name).resolve()
        self.source, self.run, self.reports = (self.root / name for name in ("source", "run", "reports"))
        self.reports.mkdir()
        for token, port, day, suffixes in (("HWLN_202601", "HWLN", "2026-01-01", ("23",)),
                                         ("KLNG_202606", "KLNG", "2026-06-01", ("11", "23"))):
            source = self.source / token
            cache = self.run / token / ".analysis_cache"
            spool = self.run / token / ".analysis_rows"
            for directory in (source, cache, spool):
                directory.mkdir(parents=True)
            digits = day.replace("-", "")
            for suffix in suffixes:
                (source / f"D&TMOK {port}_{digits}_{suffix}.xlsx").write_bytes(b"synthetic raw bytes")
            totals = {metric: 0 for metric in independent.COUNTER_NAMES}
            (cache / f"{digits}.json").write_text(json.dumps({"signature": {"port": port, "date": day},
                                                           "result": totals}), encoding="utf-8")
            (spool / f"{digits}.bin").write_bytes(b"synthetic spool bytes")
            task = independent.plan_tasks(source, self.run / token)[0]
            census = independent.RecordCensus().evidence()
            entry = independent.seal_entry({"port": port, "day": day, "pass": True, "issues": [],
                "source_files": len(suffixes), "cache": task["cache"], "spool": task["spool"],
                "totals": totals, "raw_census": census, "spool_census": census}, task)
            payload = {"checkpoint_schema": independent.CHECKPOINT_SCHEMA,
                "source_root": str(source.resolve()), "run_root": str((self.run / token).resolve()),
                "expected_source_files": len(suffixes), "expected_logical_days": 1, "completed_days": 1,
                "pass": True, "results": [entry]}
            self.path(token).write_text(json.dumps(payload), encoding="utf-8")

    def path(self, token):
        return self.reports / f"independent_raw_v155_{token}.json"

    def collect(self, expected=None):
        return collector.collect(self.source, self.run, self.reports, expected=expected)

    def mutate_report(self, action):
        path = self.path("HWLN_202601")
        report = json.loads(path.read_text(encoding="utf-8"))
        action(report)
        path.write_text(json.dumps(report), encoding="utf-8")

    def test_complete_sealed_months_are_rechecked_as_full_global_scope(self):
        result = self.collect({"source_files": 3, "logical_days": 2, "months": 2})
        self.assertTrue(result["pass"], result["issues"])
        self.assertEqual(result["status"], "COMPLETE")
        self.assertEqual(result["completed_days"], 2)
        self.assertEqual(result["expected_source_files"], 3)
        tasks = independent.plan_tasks(self.source, self.run)
        self.assertEqual(len(independent.resume_results(result, tasks, self.source, self.run)), 2)

    def test_missing_month_stays_in_progress(self):
        self.path("KLNG_202606").unlink()
        result = self.collect()
        self.assertFalse(result["pass"])
        self.assertEqual(result["status"], "IN_PROGRESS")
        self.assertEqual(result["pending_months"], ["KLNG_202606"])

    def test_atomic_partial_month_checkpoint_is_not_full_pass(self):
        path = self.path("HWLN_202601")
        report = json.loads(path.read_text(encoding="utf-8"))
        report.update({"pass": False, "results": [], "completed_days": 0})
        path.write_text(json.dumps(report), encoding="utf-8")
        result = self.collect()
        self.assertFalse(result["pass"])
        self.assertEqual(result["status"], "IN_PROGRESS")

    def test_tampered_entry_is_rejected(self):
        self.mutate_report(lambda report: report["results"][0].update(source_files=99))
        result = self.collect()
        self.assertFalse(result["pass"])
        self.assertEqual(result["status"], "FAILED")

    def test_changed_source_input_stamp_is_rejected(self):
        path = next((self.source / "HWLN_202601").glob("*.xlsx"))
        old = path.stat()
        os.utime(path, ns=(old.st_atime_ns, old.st_mtime_ns + 2_000_000_000))
        result = self.collect()
        self.assertFalse(result["pass"])
        self.assertTrue(any("inputs changed" in issue for issue in result["issues"]))

    def test_wrong_required_scope_counts_cannot_pass(self):
        result = self.collect({"source_files": 4, "logical_days": 2, "months": 2})
        self.assertFalse(result["pass"])

    def test_duplicate_sealed_day_is_rejected(self):
        self.mutate_report(lambda report: report["results"].append(report["results"][0]))
        self.assertFalse(self.collect()["pass"])

    def test_pass_report_with_incomplete_counts_is_rejected(self):
        self.mutate_report(lambda report: report.update(completed_days=0))
        self.assertFalse(self.collect()["pass"])

    def test_collector_never_modifies_month_reports_and_cancel_is_safe(self):
        before = {path: path.read_bytes() for path in self.reports.glob("*.json")}
        cancel = self.root / "cancel.flag"
        cancel.write_bytes(b"stop")
        output = self.root / "global.json"
        with contextlib.redirect_stdout(io.StringIO()):
            code = collector.main([str(self.source), str(self.run), str(self.reports),
                                   "--output", str(output), "--wait", "--cancel-file", str(cancel)])
        self.assertEqual(code, 2)
        self.assertEqual(json.loads(output.read_text(encoding="utf-8"))["status"], "PAUSED")
        self.assertEqual({path: path.read_bytes() for path in before}, before)

    def test_collector_output_cannot_replace_month_input_report(self):
        path = self.path("HWLN_202601")
        before = path.read_bytes()
        with self.assertRaises(ValueError):
            collector.main([str(self.source), str(self.run), str(self.reports), "--output", str(path)])
        self.assertEqual(path.read_bytes(), before)


if __name__ == "__main__":
    unittest.main()
