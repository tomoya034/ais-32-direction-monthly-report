from __future__ import annotations

import contextlib
import hashlib
import io
import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from scripts import verify_source_copy_manifest as verifier


class SourceCopyManifestTests(unittest.TestCase):
    """Synthetic byte fixtures exercise integrity checks; no NAS or Excel is read."""

    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory(prefix="ais-manifest-test-")
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.source = self.root / "source"
        self.copy = self.root / "copy"
        self.source.mkdir()
        self.copy.mkdir()
        self.source_file = self.source / "a.xlsx"
        self.copy_file = self.copy / "a.xlsx"
        self.source_file.write_bytes(b"data")
        self.copy_file.write_bytes(b"data")
        digest = hashlib.sha256(b"data").hexdigest().upper()
        self.manifest = {"source": str(self.source), "copy": str(self.copy), "files": [
            {"relative_path": "a.xlsx", "source_size": 4, "copy_size": 4,
             "source_sha256": digest, "copy_sha256": digest}]}
        self.manifest_path = self.root / "manifest.json"
        self.output = self.root / "report.json"
        self.checkpoint = self.output.with_suffix(".progress.json")
        self.save_manifest()

    def save_manifest(self) -> None:
        self.manifest_path.write_text(json.dumps(self.manifest), encoding="utf-8")

    def invoke(self, *extra: str) -> int:
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            return verifier.main([str(self.manifest_path), "--output", str(self.output),
                                  "--workers", "1", "--retries", "1", *extra])

    def report(self) -> dict:
        return json.loads(self.output.read_text(encoding="utf-8"))

    def change_copy(self) -> None:
        old = self.copy_file.stat()
        self.copy_file.write_bytes(b"evil")
        os.utime(self.copy_file, ns=(old.st_atime_ns, old.st_mtime_ns + 2_000_000_000))

    def test_fresh_resume_and_explicit_fresh_have_honest_counts(self) -> None:
        with mock.patch.object(verifier, "hash_file_stable", wraps=verifier.hash_file_stable) as hashing:
            self.assertEqual(self.invoke(), 0)
            self.assertEqual(hashing.call_count, 2)
            self.assertEqual(self.report()["hashed_this_run"], 2)
            self.assertEqual(self.report()["resumed_success_count"], 0)
            hashing.reset_mock()
            self.assertEqual(self.invoke(), 0)
            self.assertEqual(hashing.call_count, 0)
            self.assertEqual(self.report()["hashed_file_count"], 0)
            self.assertEqual(self.report()["task_count"], 2)
            self.assertEqual(self.report()["resumed_success_count"], 2)
            self.assertIn("not fresh", self.report()["resume_semantics"])
            hashing.reset_mock()
            self.assertEqual(self.invoke("--fresh"), 0)
            self.assertEqual(hashing.call_count, 2)
            self.assertEqual(self.report()["resumed_success_count"], 0)
        saved = json.loads(self.checkpoint.read_text(encoding="utf-8"))
        record = saved["completed"]["copy::a.xlsx"]
        self.assertEqual(record["actual_sha256"], self.manifest["files"][0]["copy_sha256"])
        self.assertIn("st_mtime_ns", record["signature"])
        self.assertTrue(record["verified_at"])

    def test_changed_manifest_same_path_cannot_reuse_success(self) -> None:
        self.assertEqual(self.invoke(), 0)
        self.manifest["files"][0]["source_sha256"] = "0" * 64
        self.manifest["files"][0]["copy_sha256"] = "0" * 64
        self.save_manifest()
        self.assertEqual(self.invoke(), 1)
        report = self.report()
        self.assertFalse(report["pass"])
        self.assertEqual(report["hashed_this_run"], 2)
        self.assertEqual(report["resumed_success_count"], 0)
        self.assertEqual(report["events"][0]["type"], "checkpoint_incompatible")

    def test_changed_root_same_manifest_path_is_rehashed(self) -> None:
        self.assertEqual(self.invoke(), 0)
        other_copy = self.root / "other_copy"
        other_copy.mkdir()
        (other_copy / "a.xlsx").write_bytes(b"evil")
        self.manifest["copy"] = str(other_copy)
        self.save_manifest()
        self.assertEqual(self.invoke(), 1)
        self.assertEqual(self.report()["resumed_success_count"], 0)

    def test_same_size_content_change_after_checkpoint_is_detected(self) -> None:
        self.assertEqual(self.invoke(), 0)
        self.change_copy()
        self.assertEqual(self.invoke(), 1)
        report = self.report()
        self.assertEqual(report["hashed_this_run"], 1)
        self.assertEqual(report["resumed_success_count"], 1)
        self.assertEqual(report["mismatches"][0]["status"], "mismatch")

    def test_old_transient_failure_is_retried_after_repair(self) -> None:
        original = verifier.hash_file_stable
        def transient(path: Path) -> tuple:
            if path == self.copy_file:
                raise OSError("synthetic temporary SMB error")
            return original(path)
        with mock.patch.object(verifier, "hash_file_stable", side_effect=transient):
            self.assertEqual(self.invoke(), 1)
        self.assertEqual(self.invoke(), 0)
        self.assertEqual(self.report()["hashed_this_run"], 1)
        self.assertEqual(self.report()["resumed_success_count"], 1)

    def test_retry_success_has_no_stale_error_and_counts_hash_attempts(self) -> None:
        original = verifier.hash_file_stable
        failed = False
        def transient_once(path: Path) -> tuple:
            nonlocal failed
            if path == self.copy_file and not failed:
                failed = True
                raise OSError("one temporary read error")
            return original(path)
        with mock.patch.object(verifier, "hash_file_stable", side_effect=transient_once), \
                mock.patch.object(verifier.time, "sleep"):
            self.assertEqual(self.invoke("--retries", "2"), 0)
        report = self.report()
        self.assertEqual(report["hash_attempt_count"], 3)
        self.assertEqual(report["hashed_this_run"], 2)
        saved = json.loads(self.checkpoint.read_text(encoding="utf-8"))
        record = saved["completed"]["copy::a.xlsx"]
        self.assertNotIn("error", record)
        self.assertEqual(len(record["retry_errors"]), 1)

    def test_legacy_null_success_is_not_reused(self) -> None:
        self.checkpoint.write_text(json.dumps({"manifest": str(self.manifest_path.resolve()),
            "scope": "both", "completed": {"source::a.xlsx": None, "copy::a.xlsx": None}}),
            encoding="utf-8")
        self.assertEqual(self.invoke(), 0)
        self.assertEqual(self.report()["hashed_this_run"], 2)
        self.assertEqual(self.report()["events"][0]["type"], "checkpoint_incompatible")

    def test_corrupt_checkpoint_rehashes_and_records_event(self) -> None:
        self.checkpoint.write_text('{"completed":', encoding="utf-8")
        self.assertEqual(self.invoke(), 0)
        self.assertEqual(self.report()["hashed_this_run"], 2)
        self.assertEqual(self.report()["events"][0]["type"], "checkpoint_unreadable")

    def test_tampered_success_record_is_not_trusted(self) -> None:
        self.assertEqual(self.invoke(), 0)
        saved = json.loads(self.checkpoint.read_text(encoding="utf-8"))
        saved["completed"]["copy::a.xlsx"]["actual_sha256"] = "0" * 64
        self.checkpoint.write_text(json.dumps(saved), encoding="utf-8")
        self.assertEqual(self.invoke(), 0)
        self.assertEqual(self.report()["hashed_this_run"], 1)

    def test_content_change_during_hash_is_rejected(self) -> None:
        original_sha256 = hashlib.sha256
        digest = original_sha256()
        file = self.copy_file
        class ChangingDigest:
            changed = False

            def update(self, chunk: bytes) -> None:
                digest.update(chunk)
                if not self.changed:
                    self.changed = True
                    with file.open("ab") as handle:
                        handle.write(b"x")
            def hexdigest(self) -> str:
                return digest.hexdigest()
        with mock.patch.object(verifier.hashlib, "sha256", return_value=ChangingDigest()):
            with self.assertRaises(verifier.FileChangedError):
                verifier.hash_file_stable(file)

    def test_file_identity_change_between_stat_and_open_is_rejected(self) -> None:
        original_signature = verifier.signature
        calls = 0
        def different_identity(value: os.stat_result) -> dict:
            nonlocal calls
            calls += 1
            result = original_signature(value)
            if calls == 2:
                result["st_ino"] += 1
            return result
        with mock.patch.object(verifier, "signature", side_effect=different_identity):
            with self.assertRaises(verifier.FileChangedError):
                verifier.hash_file_stable(self.copy_file)

    def test_short_read_is_rejected_even_when_stat_is_unchanged(self) -> None:
        with self.copy_file.open("rb") as real_handle:
            fake = mock.MagicMock()
            fake.__enter__.return_value = fake
            fake.fileno.return_value = real_handle.fileno()
            fake.read.return_value = b""
            with mock.patch.object(Path, "open", return_value=fake):
                with self.assertRaises(verifier.FileChangedError):
                    verifier.hash_file_stable(self.copy_file)

    def test_change_after_success_before_final_validation_cannot_pass(self) -> None:
        original = verifier.atomic_json
        changed = False
        def write_and_change(path: Path, value: object) -> None:
            nonlocal changed
            original(path, value)
            if path == self.checkpoint and not changed:
                changed = True
                self.change_copy()
        with mock.patch.object(verifier, "atomic_json", side_effect=write_and_change):
            self.assertEqual(self.invoke(), 1)
        self.assertIn("after verification", self.report()["mismatches"][0]["error"])

    def test_file_added_after_start_inventory_cannot_pass(self) -> None:
        original = verifier.hash_file_stable
        def hash_and_add(path: Path) -> tuple:
            result = original(path)
            (self.copy / "extra.xlsx").write_bytes(b"extra")
            return result
        with mock.patch.object(verifier, "hash_file_stable", side_effect=hash_and_add):
            self.assertEqual(self.invoke(), 1)
        report = self.report()
        self.assertTrue(report["roots"]["copy"]["inventory_changed"])
        self.assertEqual(report["roots"]["copy"]["extra"], ["extra.xlsx"])

    def test_missing_source_is_not_a_passing_inventory(self) -> None:
        self.copy_file.unlink()
        self.assertEqual(self.invoke(), 1)
        self.assertEqual(self.report()["roots"]["copy"]["missing"], ["a.xlsx"])
        self.assertEqual(self.report()["attempted_this_run"], 2)
        self.assertEqual(self.report()["hashed_this_run"], 1)

    def test_file_changed_during_ending_inventory_cannot_pass(self) -> None:
        original = verifier.inventory
        calls = 0
        def inventory_and_change(root: Path, expected: set[str]) -> dict:
            nonlocal calls
            calls += 1
            result = original(root, expected)
            if calls == 4:
                self.change_copy()
            return result
        with mock.patch.object(verifier, "inventory", side_effect=inventory_and_change):
            self.assertEqual(self.invoke("--fresh"), 1)
        report = self.report()
        self.assertEqual(report["hashed_this_run"], 2)
        self.assertIn("after verification", report["mismatches"][0]["error"])
        self.assertIn("not an atomic", report["snapshot_semantics"])

    def test_both_scope_rejects_different_source_copy_expected_content(self) -> None:
        self.copy_file.write_bytes(b"evil")
        self.manifest["files"][0]["copy_sha256"] = hashlib.sha256(b"evil").hexdigest().upper()
        self.manifest["files"][0]["match"] = False
        self.save_manifest()
        with mock.patch.object(verifier, "hash_file_stable") as hashing:
            self.assertEqual(self.invoke("--fresh"), 2)
            hashing.assert_not_called()
        self.assertFalse(self.output.exists())
        self.assertEqual(self.copy_file.read_bytes(), b"evil")

    def test_both_scope_rejects_different_expected_sizes(self) -> None:
        self.manifest["files"][0]["copy_size"] += 1
        self.save_manifest()
        self.assertEqual(self.invoke("--fresh"), 2)
        self.assertFalse(self.output.exists())

    def test_inventory_scan_error_cannot_become_pass(self) -> None:
        original = verifier.os.walk
        def scan(root: Path, onerror=None):
            if Path(root) == self.copy:
                onerror(PermissionError("synthetic inaccessible directory"))
            return original(root, onerror=onerror)
        with mock.patch.object(verifier.os, "walk", side_effect=scan):
            self.assertEqual(self.invoke(), 1)
        self.assertTrue(self.report()["roots"]["copy"]["errors"])

    def test_manifest_changed_during_run_cannot_pass(self) -> None:
        original = verifier.hash_file_stable
        def hash_and_change_manifest(path: Path) -> tuple:
            result = original(path)
            self.manifest["changed"] = True
            self.save_manifest()
            return result
        with mock.patch.object(verifier, "hash_file_stable", side_effect=hash_and_change_manifest):
            self.assertEqual(self.invoke(), 1)
        self.assertFalse(self.report()["manifest_unchanged"])

    def test_atomic_write_failure_preserves_previous_json(self) -> None:
        self.output.write_text('{"historical":true}', encoding="utf-8")
        with mock.patch.object(verifier.os, "replace", side_effect=OSError("synthetic interruption")):
            with self.assertRaises(OSError):
                verifier.atomic_json(self.output, {"new": True})
        self.assertEqual(self.report(), {"historical": True})
        self.assertEqual(list(self.root.glob("*.tmp")), [])

    def test_report_write_failure_returns_nonzero(self) -> None:
        original = verifier.atomic_json
        def report_failure(path: Path, value: object) -> None:
            if path == self.output:
                raise OSError("synthetic report write failure")
            original(path, value)
        with mock.patch.object(verifier, "atomic_json", side_effect=report_failure):
            self.assertEqual(self.invoke(), 2)
        self.assertFalse(self.output.exists())

    def test_same_checkpoint_cannot_be_used_by_two_instances(self) -> None:
        self.output.write_text('{"historical":true}', encoding="utf-8")
        with verifier.exclusive_lock(self.checkpoint):
            self.assertEqual(self.invoke(), 2)
        self.assertEqual(self.report(), {"historical": True})
        # The remaining lock file is harmless after the OS releases its lock.
        self.assertEqual(self.invoke(), 0)

    def test_duplicate_relative_paths_are_rejected(self) -> None:
        self.manifest["files"].append(dict(self.manifest["files"][0]))
        self.save_manifest()
        self.assertEqual(self.invoke(), 2)
        self.assertFalse(self.output.exists())

    def test_parent_relative_path_is_rejected(self) -> None:
        self.manifest["files"][0]["relative_path"] = "../a.xlsx"
        self.save_manifest()
        self.assertEqual(self.invoke(), 2)
        self.assertFalse(self.output.exists())

    def test_report_cannot_overwrite_source_or_copy_data(self) -> None:
        for protected in (self.source_file, self.copy_file):
            with self.subTest(protected=protected):
                before = protected.read_bytes()
                with mock.patch.object(verifier, "hash_file_stable") as hashing:
                    self.assertEqual(self.invoke("--fresh", "--output", str(protected)), 2)
                    hashing.assert_not_called()
                self.assertEqual(protected.read_bytes(), before)

    def test_report_inside_unselected_source_root_is_rejected_before_writes(self) -> None:
        target = self.source / "audit.json"
        self.assertEqual(self.invoke("--scope", "copy", "--output", str(target)), 2)
        self.assertFalse(target.exists())
        self.assertFalse(target.with_suffix(".progress.json").exists())
        self.assertEqual(self.source_file.read_bytes(), b"data")

    def test_checkpoint_inside_copy_root_is_rejected_before_writes(self) -> None:
        target = self.copy / "new.progress.json"
        self.assertEqual(self.invoke("--checkpoint", str(target)), 2)
        self.assertFalse(target.exists())
        self.assertFalse(target.with_name(target.name + ".lock").exists())
        self.assertEqual(self.copy_file.read_bytes(), b"data")

    def test_report_and_checkpoint_cannot_overwrite_manifest(self) -> None:
        before = self.manifest_path.read_bytes()
        self.assertEqual(self.invoke("--output", str(self.manifest_path)), 2)
        self.assertEqual(self.manifest_path.read_bytes(), before)
        self.assertEqual(self.invoke("--checkpoint", str(self.manifest_path)), 2)
        self.assertEqual(self.manifest_path.read_bytes(), before)

    def test_checkpoint_lock_cannot_overwrite_manifest(self) -> None:
        old_manifest = self.manifest_path
        self.manifest_path = self.root / "audit.progress.lock"
        self.manifest_path.write_bytes(old_manifest.read_bytes())
        before = self.manifest_path.read_bytes()
        self.assertEqual(self.invoke("--checkpoint", str(self.root / "audit.progress")), 2)
        self.assertEqual(self.manifest_path.read_bytes(), before)

    def test_report_symlink_in_copy_root_cannot_place_temporary_files_there(self) -> None:
        target = self.copy / "report.json"
        external = self.root / "external.json"
        # Model final symlink resolution without requiring Windows symlink rights.
        original = Path.resolve
        def resolved(path: Path, *args, **kwargs) -> Path:
            if path == target:
                return external
            return original(path, *args, **kwargs)
        with mock.patch.object(Path, "resolve", resolved):
            self.assertEqual(self.invoke("--output", str(target)), 2)
        self.assertFalse(target.exists())
        self.assertFalse(external.exists())


if __name__ == "__main__":
    unittest.main()
