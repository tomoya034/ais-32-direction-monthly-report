import contextlib
import io
import json
import os
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from scripts import run_packaged_full_regression as packaged


class PackagedRunnerTests(unittest.TestCase):
    def test_source_overlap_rejected_before_writing(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            executable = root / "tool.exe"
            executable.write_bytes(b"fixture")
            source = root / "source"
            source.mkdir()
            for output in (source, source / "child", root):
                with self.subTest(output=output), self.assertRaises(ValueError):
                    packaged.guard_roots(source, output, executable)
            self.assertEqual(list(source.iterdir()), [])

    def test_executable_container_cannot_be_output(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            executable = root / "package" / "tool.exe"
            executable.parent.mkdir()
            executable.write_bytes(b"fixture")
            with self.assertRaises(ValueError):
                packaged.guard_roots(root / "source", executable.parent, executable)

    def test_empty_plan_cannot_pass_or_write_report(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            executable = root / "tool.exe"
            executable.write_bytes(b"fixture")
            output = root / "outputs"
            with mock.patch.object(packaged, "discover_jobs", return_value=[]):
                with contextlib.redirect_stdout(io.StringIO()), self.assertRaises(ValueError):
                    packaged.main([str(root / "source"), str(output), "--executable", str(executable)])
            self.assertFalse(output.exists())

    def test_existing_evidence_or_cache_cannot_be_overwritten(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            executable = root / "tool.exe"
            executable.write_bytes(b"fixture")
            output = root / "outputs"
            output.mkdir()
            evidence = output / "run_report.json"
            evidence.write_bytes(b"preserve prior checkpoint")
            with self.assertRaises(ValueError):
                packaged.guard_roots(root / "source", output, executable)
            self.assertEqual(evidence.read_bytes(), b"preserve prior checkpoint")

    def test_empty_output_directory_is_allowed_but_file_is_not(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            executable = root / "tool.exe"
            executable.write_bytes(b"fixture")
            output = root / "outputs"
            output.mkdir()
            packaged.guard_roots(root / "source", output, executable)
            output.rmdir()
            output.write_bytes(b"not a directory")
            with self.assertRaises(ValueError):
                packaged.guard_roots(root / "source", output, executable)
            self.assertEqual(output.read_bytes(), b"not a directory")


class WindowsOwnershipTests(unittest.TestCase):
    def setUp(self):
        self.output = Path(tempfile.gettempdir()) / "ais-owned-output"
        self.exe = Path(tempfile.gettempdir()) / "tool.exe"
        self.args = ["python.exe", "-m", "scripts.run_packaged_full_regression",
                     "source", str(self.output), "--executable", str(self.exe), "--prepare-resume"]
        self.current = self.process(200, 100, self.args)
        self.parent = self.process(100, 50, self.args)

    def process(self, pid, parent, args, *, name="python.exe", image=None):
        return {"ProcessId": pid, "ParentProcessId": parent, "Name": name,
                "ExecutablePath": image or str(Path(tempfile.gettempdir()) / name),
                "CommandLine": json.dumps(args)}

    def check(self, others):
        with mock.patch.object(packaged.os, "getpid", return_value=200), \
                mock.patch.object(packaged, "windows_argv", side_effect=json.loads):
            packaged.reject_live_writers([self.current, *others], self.output, self.exe)

    def test_identical_direct_venv_launcher_is_our_own_process(self):
        self.check([self.parent])

    def test_another_identical_wrapper_not_our_parent_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "PID 300"):
            self.check([self.parent, self.process(300, 50, self.args)])

    def test_direct_parent_with_different_cli_is_not_exempt(self):
        args = self.args[:-1] + ["--resume"]
        with self.assertRaisesRegex(ValueError, "PID 100"):
            self.check([self.process(100, 50, args)])

    def test_embedded_module_name_in_powershell_is_not_a_wrapper(self):
        self.check([self.parent, self.process(50, 1, ["powershell.exe", "-Command",
                    "python -m scripts.run_packaged_full_regression " + str(self.output)], name="powershell.exe")])

    def test_actual_frozen_producer_is_always_rejected(self):
        with self.assertRaisesRegex(ValueError, "PID 400"):
            self.check([self.process(400, 1, [], name="tool.exe", image=str(self.exe))])

    def test_missing_current_or_uninspectable_producer_fails_closed(self):
        with mock.patch.object(packaged.os, "getpid", return_value=999):
            with self.assertRaisesRegex(ValueError, "Current process missing"):
                packaged.reject_live_writers([self.current], self.output, self.exe)
        hidden = {"ProcessId": 400, "Name": "tool.exe", "ExecutablePath": None}
        with self.assertRaisesRegex(ValueError, "Cannot inspect"):
            self.check([hidden])

    @unittest.skipUnless(os.name == "nt", "Windows argv API")
    def test_native_windows_argv_preserves_spaces_quotes_and_unicode(self):
        args = [r"C:\task path\python.exe", "-m", "scripts.run_packaged_full_regression",
                r"B:\input 花蓮", str(self.output), "--prepare-resume"]
        self.assertEqual(packaged.windows_argv(packaged.subprocess.list2cmdline(args)), args)


class PackagedResumeTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix="ais-packaged-resume-")
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.source, self.output = self.root / "source", self.root / "output"
        self.source.mkdir()
        self.output.mkdir()
        self.exe = self.root / "tool.exe"
        self.exe.write_bytes(b"frozen fixture executable")
        self.planned = []
        for month in (1, 2):
            folder = self.source / f"HWLN_2026{month:02d}"
            folder.mkdir()
            filename = f"D&TMOK HWLN_2026{month:02d}01_23.xlsx"
            (folder / filename).write_bytes(b"protected source " + bytes([month]))
            self.planned.append({"folder": str(folder), "port": "HWLN", "year": 2026, "month": month,
                "physical_files": 1, "tolerant_logical_days": 1, "logical_days": [f"2026-{month:02d}-01"],
                "source_files": [filename], "token": f"HWLN_2026{month:02d}"})
        self.completed = self.create_month(self.planned[0])
        self.partial = self.output / "HWLN_202602"
        self.partial.mkdir()
        (self.partial / "partial.bin").write_bytes(b"preserve entire partial month")
        self.report = {"scope": "FULL", "source_root": str(self.source), "started_at": "original-start",
            "status": "INTERRUPTED", "planned_jobs": self.planned, "expected_job_count": 2,
            "jobs": [self.completed], "active_job": "HWLN_202602",
            "executable": str(self.exe), "executable_sha256": packaged.executable_hash(self.exe)}
        self.save()
        self.inactive = mock.patch.object(packaged, "assert_inactive")
        self.inactive.start()
        self.addCleanup(self.inactive.stop)

    def save(self):
        (self.output / "run_report.json").write_text(json.dumps(self.report), encoding="utf-8")

    def create_month(self, job):
        modern, delivery, files, cache, spool = packaged.job_paths(job, self.output)
        for path in files:
            path.parent.mkdir(parents=True, exist_ok=True)
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("[Content_Types].xml", "<Types/>")
        cache.mkdir()
        spool.mkdir()
        day = str(job["logical_days"][0]).replace("-", "")
        spool_path = spool / (day + ".bin")
        spool_path.write_bytes(b"normalized source rows")
        source = Path(job["folder"]) / job["source_files"][0]
        payload = {"signature": {"date": job["logical_days"][0], "port": "HWLN", "year": 2026,
            "month": job["month"], "sources": [{"path": str(source), "size": source.stat().st_size,
            "modified_ns": source.stat().st_mtime_ns}]}, "result": {"day": job["logical_days"][0],
            "source_fragments": [str(source)], "fragment_info": [{"path": str(source),
            "sha256": packaged.executable_hash(source)}], "spool_sha256": packaged.executable_hash(spool_path)}}
        (cache / (day + ".json")).write_text(json.dumps(payload), encoding="utf-8")
        return {**job, "modern": str(modern), "delivery": str(delivery), "status": "PASS",
            "exit_code": 0, "observed_xlsx_count": 6, "current_parser_logical_days": 1}

    def prepare(self):
        return packaged.resume_checkpoint(self.source, self.output, self.exe, self.planned, prepare=True)

    def test_preparation_preserves_original_report_and_archives_whole_partial_month(self):
        original = (self.output / "run_report.json").read_bytes()
        report = self.prepare()
        archive = Path(report["resume_events"][-1]["archive"])
        self.assertEqual((archive / "run_report.json").read_bytes(), original)
        self.assertEqual((archive / "HWLN_202602/partial.bin").read_bytes(), b"preserve entire partial month")
        self.assertFalse(self.partial.exists())
        self.assertEqual(report["started_at"], "original-start")
        self.assertEqual(report["status"], "PAUSED_READY")
        self.assertEqual(len(report["jobs"][0]["resume_file_sha256"]), 8)
        checked = packaged.resume_checkpoint(self.source, self.output, self.exe, self.planned, prepare=False)
        self.assertEqual(len(checked["jobs"]), 1)

    def test_actual_resume_skips_completed_month_and_restarts_missing_month_from_raw(self):
        self.prepare()
        def run(command, **kwargs):
            self.assertIn("--input", command)
            self.create_month(self.planned[1])
            return mock.Mock(returncode=0)
        with mock.patch.object(packaged, "discover_jobs", return_value=self.planned), \
                mock.patch.object(packaged.subprocess, "run", side_effect=run) as executed, \
                contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(packaged.main([str(self.source), str(self.output), "--executable", str(self.exe), "--resume"]), 0)
        self.assertEqual(executed.call_count, 1)
        report = json.loads((self.output / "run_report.json").read_text(encoding="utf-8"))
        self.assertEqual(report["started_at"], "original-start")
        self.assertEqual(len(report["jobs"]), 2)

    def test_plan_or_executable_change_is_rejected_before_archiving(self):
        before = (self.output / "run_report.json").read_bytes()
        self.report["planned_jobs"] = self.planned[:1]
        self.save()
        with self.assertRaisesRegex(ValueError, "plan/hash"):
            self.prepare()
        self.assertTrue(self.partial.exists())
        self.report["planned_jobs"] = self.planned
        self.save()
        self.exe.write_bytes(b"different executable")
        with self.assertRaisesRegex(ValueError, "executable identity"):
            self.prepare()
        self.assertTrue(self.partial.exists())

    def test_resume_requires_explicit_preparation(self):
        with self.assertRaisesRegex(ValueError, "prepare-resume"):
            packaged.resume_checkpoint(self.source, self.output, self.exe, self.planned, prepare=False)

    def test_missing_output_or_spool_cannot_retain_completed_month(self):
        files = packaged.job_paths(self.planned[0], self.output)[2]
        files[0].unlink()
        with self.assertRaisesRegex(ValueError, "final XLSX"):
            self.prepare()
        self.assertTrue(self.partial.exists())

    def test_source_hash_change_is_detected_even_when_size_and_mtime_are_restored(self):
        source = Path(self.planned[0]["folder"]) / self.planned[0]["source_files"][0]
        timestamp = source.stat().st_mtime_ns
        source.write_bytes(b"protected source " + bytes([9]))
        os.utime(source, ns=(timestamp, timestamp))
        with self.assertRaisesRegex(ValueError, "source content/signature"):
            self.prepare()

    def test_noncanonical_stored_artifact_path_is_rejected(self):
        self.report["jobs"][0]["modern"] = str(self.output / "../output" / "modern.xlsx")
        self.save()
        with self.assertRaisesRegex(ValueError, "Noncanonical"):
            self.prepare()

    def test_changed_cached_data_after_checkpoint_fails_stored_fingerprint(self):
        self.prepare()
        cache = packaged.job_paths(self.planned[0], self.output)[3] / "20260101.json"
        payload = json.loads(cache.read_text())
        payload["unexpected_data_edit"] = 999
        cache.write_text(json.dumps(payload))
        with self.assertRaisesRegex(ValueError, "fingerprints changed"):
            packaged.resume_checkpoint(self.source, self.output, self.exe, self.planned, prepare=False)

    def test_live_writer_refusal_leaves_report_and_partial_evidence_unchanged(self):
        before = (self.output / "run_report.json").read_bytes()
        with mock.patch.object(packaged, "assert_inactive", side_effect=ValueError("live writer")):
            with self.assertRaisesRegex(ValueError, "live writer"):
                self.prepare()
        self.assertEqual((self.output / "run_report.json").read_bytes(), before)
        self.assertTrue(self.partial.exists())

    def test_failed_month_is_archived_and_never_reused_as_completed(self):
        self.report["jobs"][0]["status"] = "FAIL"
        self.save()
        report = self.prepare()
        self.assertEqual(report["jobs"], [])
        self.assertFalse((self.output / "HWLN_202601").exists())

    def test_pause_before_next_month_retains_completed_without_launch(self):
        self.prepare()
        flag = self.root.parent / (self.root.name + ".quota.pause")
        self.addCleanup(lambda: flag.unlink(missing_ok=True))
        flag.write_text("User observed remaining5", encoding="utf-8")
        with mock.patch.object(packaged, "discover_jobs", return_value=self.planned), \
                mock.patch.object(packaged.subprocess, "run") as execute, \
                contextlib.redirect_stdout(io.StringIO()):
            code = packaged.main([str(self.source), str(self.output), "--executable", str(self.exe),
                                  "--resume", "--pause-flag", str(flag)])
        self.assertEqual(code, 2)
        execute.assert_not_called()
        report = json.loads((self.output / "run_report.json").read_text(encoding="utf-8"))
        self.assertEqual(report["status"], "PAUSED_BY_USER_QUOTA")
        self.assertEqual(report["pause_stage"], "BEFORE_NEXT_MONTH")
        self.assertEqual(len(report["jobs"]), 1)
        self.assertEqual(report["jobs"][0]["status"], "PASS")

    def test_pause_after_stopped_child_never_counts_data_fail_or_starts_next_job(self):
        self.prepare()
        flag = self.root.parent / (self.root.name + ".quota.pause")
        self.addCleanup(lambda: flag.unlink(missing_ok=True))
        def run(command, **kwargs):
            flag.write_text("User observed remaining5", encoding="utf-8")
            modern = packaged.job_paths(self.planned[1], self.output)[0]
            modern.parent.mkdir(parents=True, exist_ok=True)
            (modern.parent / "own_partial.bin").write_bytes(b"preserve partial")
            return mock.Mock(returncode=4294967295)
        with mock.patch.object(packaged, "discover_jobs", return_value=self.planned), \
                mock.patch.object(packaged.subprocess, "run", side_effect=run) as execute, \
                contextlib.redirect_stdout(io.StringIO()):
            code = packaged.main([str(self.source), str(self.output), "--executable", str(self.exe),
                                  "--resume", "--pause-flag", str(flag)])
        self.assertEqual(code, 2)
        self.assertEqual(execute.call_count, 1)
        report = json.loads((self.output / "run_report.json").read_text(encoding="utf-8"))
        self.assertEqual(report["status"], "PAUSED_BY_USER_QUOTA")
        self.assertEqual(report["pause_stage"], "AFTER_ACTIVE_CHILD")
        self.assertEqual(len(report["jobs"]), 1)
        self.assertEqual(report["active_job"], "HWLN_202602")
        self.assertEqual((self.output / "HWLN_202602/own_partial.bin").read_bytes(), b"preserve partial")

    def test_pause_flag_cannot_be_in_protected_input_output_exe_tree(self):
        for flag in (self.source / "flag", self.output / "flag", self.exe.parent / "flag"):
            with self.subTest(flag=flag), self.assertRaisesRegex(ValueError, "outside"):
                packaged.main([str(self.source), str(self.output), "--executable", str(self.exe),
                               "--pause-flag", str(flag), "--resume"])


if __name__ == "__main__":
    unittest.main()
