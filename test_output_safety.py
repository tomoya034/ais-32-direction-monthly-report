from __future__ import annotations

import hashlib
import ast
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import unittest
from pathlib import Path
from unittest import mock
from types import SimpleNamespace

import openpyxl

import ais_monthly_app as app
import test_v15_contracts as fixtures


class OutputSafetyTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory(prefix="ais32-output-safety-")
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.source = self.root / "D&TMOK KLNG_20260601_any.xlsx"
        fixtures.FragmentCatalogTests._write_records(
            self.source,
            [(1, 0, "A", 1, 50.0), (1, 1, "A", 1, 49.0), (1, 2, "B", 1, 48.0),
             (13, 0, "A", 3, 70.0), (13, 1, "A", 3, 69.0), (13, 2, "B", 3, 68.0)],
        )
        self.config = app.AppConfig(
            self.root, self.root / "analysis.xlsx", "KLNG", 2026, 6,
            delivery_dir=self.root / "delivery", overwrite=True, workers=1,
        )

    def hashes(self) -> dict[Path, str]:
        return {path: app._file_sha256(path) for path in
                (self.config.output_path, *app.derive_delivery_paths(self.config).all_files())}

    def gui_callback(self, method_name: str, config: app.AppConfig):
        tree = ast.parse(Path(app.__file__).read_text(encoding="utf-8"))
        application = next(node for node in ast.walk(tree) if isinstance(node, ast.ClassDef) and node.name == "Application")
        method = next(node for node in application.body if isinstance(node, ast.FunctionDef) and node.name == method_name)
        thread_factory = mock.Mock()
        messagebox, filedialog = mock.Mock(), mock.Mock()
        namespace = {**app.__dict__, "messagebox": messagebox, "filedialog": filedialog,
                     "threading": SimpleNamespace(Thread=thread_factory)}
        exec(compile(ast.Module(body=[method], type_ignores=[]), str(app.__file__), "exec"), namespace)
        fake = SimpleNamespace(
            worker=None, build_config=lambda: config, cancel_event=threading.Event(),
            start_button=mock.Mock(), finalize_button=mock.Mock(), cancel_button=mock.Mock(),
            open_button=mock.Mock(), progress={}, status_var=mock.Mock(), progress_text_var=mock.Mock(),
            append_log=mock.Mock(), events=app.queue.Queue(),
        )
        return namespace[method_name], fake, messagebox, filedialog, thread_factory

    def test_modern_historical_path_collision_rejected_before_writing(self) -> None:
        self.config.output_path = app.derive_delivery_paths(self.config).integrated_summary
        with self.assertRaisesRegex(ValueError, "路徑碰撞"):
            app.run_pipeline(self.config)
        self.assertFalse(self.config.output_path.exists())

    def test_raw_source_cannot_be_an_output(self) -> None:
        before = app._file_sha256(self.source)
        self.config.output_path = self.source
        with self.assertRaisesRegex(ValueError, "原始 AIS"):
            app.run_pipeline(self.config)
        self.assertEqual(app._file_sha256(self.source), before)

    def test_delivery_cannot_use_normalized_storage(self) -> None:
        self.config.delivery_dir = app._normalized_spool_directory(self.config)
        with self.assertRaisesRegex(ValueError, "cache／normalized spool"):
            self.config.validate()

    def test_building_path_cannot_delete_an_existing_raw_source(self) -> None:
        self.config.output_path = self.root / "D&TMOK KLNG_20260601_new.xlsx"
        protected = self.config.output_path.with_name(self.config.output_path.stem + ".building.xlsx")
        shutil.copyfile(self.source, protected)
        before = app._file_sha256(protected)
        with self.assertRaisesRegex(ValueError, "原始 AIS"):
            app.run_pipeline(self.config)
        self.assertEqual(app._file_sha256(protected), before)

    def test_modern_cannot_use_a_historical_building_path(self) -> None:
        target = app.derive_delivery_paths(self.config).integrated_summary
        self.config.output_path = target.with_name(target.stem + ".building.xlsx")
        with self.assertRaisesRegex(ValueError, "temporary／rollback 路徑碰撞"):
            self.config.validate()

    def test_parallel_different_modern_targets_same_delivery_rejected(self) -> None:
        results, _ = app.process_month(self.config)
        snapshot = app.build_default_decision_snapshot(self.config, results)
        entered = threading.Event()
        release = threading.Event()
        failures = []

        def callback(payload):
            if payload.get("kind") == "delivery_file_done" and payload.get("position") == 1:
                entered.set()
                if not release.wait(15):
                    raise RuntimeError("test writer release timed out")

        def writer():
            try:
                app.write_delivery_workbooks(self.config, results, snapshot, callback=callback)
            except BaseException as error:
                failures.append(error)

        thread = threading.Thread(target=writer)
        thread.start()
        try:
            self.assertTrue(entered.wait(15))
            other = app.AppConfig(
                self.root, self.root / "other.xlsx", "KLNG", 2026, 6,
                delivery_dir=self.config.delivery_dir, overwrite=True, workers=1,
            )
            with self.assertRaisesRegex(RuntimeError, "另一個工作"):
                app.write_delivery_workbooks(other, results, snapshot)
        finally:
            release.set()
            thread.join(15)
        self.assertFalse(thread.is_alive())
        self.assertEqual(failures, [])
        paths = app.derive_delivery_paths(self.config)
        for path, sheet_name in [(paths.period_b_full, "總表"), (paths.period_b_summary, None)]:
            workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
            try:
                sheet = workbook[sheet_name] if sheet_name else workbook.worksheets[0]
                self.assertEqual(sheet["B2"].value, 70)
            finally:
                workbook.close()

    def test_cross_process_same_output_is_rejected(self) -> None:
        child_code = (
            "from pathlib import Path; import ais_monthly_app as a; import sys; "
            "c=a.AppConfig(Path(sys.argv[1]),Path(sys.argv[2]),'KLNG',2026,6,"
            "delivery_dir=Path(sys.argv[3]),overwrite=True,workers=1); a.process_month(c)"
        )
        with app._exclusive_job(self.config):
            child = subprocess.run(
                [sys.executable, "-c", child_code, str(self.root),
                 str(self.config.output_path), str(self.config.delivery_dir)],
                cwd=Path(__file__).parent, capture_output=True, timeout=15,
                env={**os.environ, "PYTHONIOENCODING": "utf-8"},
            )
        self.assertNotEqual(child.returncode, 0)
        self.assertIn("另一個工作", child.stderr.decode("utf-8"))
        self.assertFalse(self.config.output_path.exists())
        self.assertFalse(any(self.root.rglob("*.lock")))

    def test_cross_job_final_target_cannot_use_another_jobs_temporary_path(self) -> None:
        other = app.AppConfig(
            self.root, self.config.output_path.with_name(self.config.output_path.stem + ".building.xlsx"),
            "KLNG", 2026, 6, delivery_dir=self.root / "other_delivery", overwrite=True, workers=1,
        )
        with app._exclusive_job(self.config):
            error = []

            def competing_writer():
                try:
                    app.process_month(other)
                except RuntimeError as failure:
                    error.append(str(failure))

            thread = threading.Thread(target=competing_writer)
            thread.start()
            thread.join(15)
            self.assertFalse(thread.is_alive())
            self.assertEqual(len(error), 1)
            self.assertIn("另一個工作", error[0])
        self.assertFalse(other.output_path.exists())

    def test_stale_lock_is_preserved_for_manual_review(self) -> None:
        key = app._canonical_path(self.config.output_path)
        lock = Path(tempfile.gettempdir()) / "AIS32_job_locks" / (hashlib.sha256(key.encode()).hexdigest() + ".lock")
        lock.parent.mkdir(parents=True, exist_ok=True)
        payload = json.dumps({"pid": 99999999, "token": "stale-evidence", "resource": key})
        lock.write_text(payload, encoding="utf-8")
        try:
            with self.assertRaisesRegex(RuntimeError, "上次工作留下 lock"):
                app.process_month(self.config)
            self.assertEqual(lock.read_text(encoding="utf-8"), payload)
        finally:
            lock.unlink()

    def test_historical_build_failure_preserves_all_six_previous_outputs(self) -> None:
        app.run_pipeline(self.config)
        before = self.hashes()
        with mock.patch.object(app, "_write_summary_workbook", side_effect=OSError("injected build failure")):
            with self.assertRaisesRegex(OSError, "injected build failure"):
                app.run_pipeline(self.config)
        self.assertEqual(self.hashes(), before)
        self.assertFalse(any(self.root.glob(".ais32-stage-*")))

    def test_six_file_promotion_failure_restores_every_previous_output(self) -> None:
        app.run_pipeline(self.config)
        before = self.hashes()
        real_replace = os.replace
        promotions = 0

        def fail_third_promotion(source, destination):
            nonlocal promotions
            if Path(destination) in before and ".ais32-stage-" in str(source):
                promotions += 1
                if promotions == 3:
                    raise OSError("injected six-file promotion failure")
            return real_replace(source, destination)

        with mock.patch.object(app.os, "replace", side_effect=fail_third_promotion):
            with self.assertRaisesRegex(OSError, "injected six-file promotion failure"):
                app.run_pipeline(self.config)
        self.assertEqual(self.hashes(), before)
        self.assertFalse(any(self.root.rglob("*.rollback")))

    def test_late_existing_output_is_not_overwritten_without_permission(self) -> None:
        self.config.overwrite = False
        sentinel = b"another-tool-created-this-after-preflight"

        def callback(payload):
            if payload.get("kind") == "delivery_write_done":
                self.config.output_path.write_bytes(sentinel)

        with self.assertRaises(FileExistsError):
            app.run_pipeline(self.config, callback=callback)
        self.assertEqual(self.config.output_path.read_bytes(), sentinel)
        self.assertFalse(any(path.exists() for path in app.derive_delivery_paths(self.config).all_files()))

    def test_cancel_after_staging_preserves_all_six_previous_outputs(self) -> None:
        app.run_pipeline(self.config)
        before = self.hashes()
        cancel = threading.Event()

        def callback(payload):
            if payload.get("kind") == "delivery_write_done":
                cancel.set()

        with self.assertRaises(app.CancelledError):
            app.run_pipeline(self.config, callback=callback, cancel_event=cancel)
        self.assertEqual(self.hashes(), before)

    def test_late_cancel_of_five_file_finalize_preserves_previous_historical_files(self) -> None:
        results, _ = app.process_month(self.config)
        snapshot = app.build_default_decision_snapshot(self.config, results)
        paths = app.write_delivery_workbooks(self.config, results, snapshot)
        before = {path: app._file_sha256(path) for path in paths.all_files()}
        cancel = threading.Event()

        def callback(payload):
            if payload.get("kind") == "delivery_file_done" and payload.get("position") == 5:
                cancel.set()

        with self.assertRaises(app.CancelledError):
            app.write_delivery_workbooks(self.config, results, snapshot, callback=callback, cancel_event=cancel)
        self.assertEqual(before, {path: app._file_sha256(path) for path in paths.all_files()})

    def test_standalone_five_file_writer_preserves_late_created_output(self) -> None:
        results, _ = app.process_month(self.config)
        snapshot = app.build_default_decision_snapshot(self.config, results)
        self.config.overwrite = False
        paths = app.derive_delivery_paths(self.config)
        sentinel = b"created-by-another-tool-after-five-file-preflight"

        def callback(payload):
            if payload.get("kind") == "delivery_file_done" and payload.get("position") == 5:
                paths.integrated_summary.write_bytes(sentinel)

        with self.assertRaises(FileExistsError):
            app.write_delivery_workbooks(self.config, results, snapshot, callback=callback)
        self.assertEqual(paths.integrated_summary.read_bytes(), sentinel)
        self.assertFalse(any(p.exists() for p in paths.all_files() if p != paths.integrated_summary))

    def test_no_overwrite_rejects_a_file_created_during_atomic_promotion(self) -> None:
        self.config.overwrite = False
        sentinel = b"created-between-final-check-and-rename"
        operation = "rename" if os.name == "nt" else "link"
        real_operation = getattr(os, operation)

        def concurrent_create(source, destination):
            if Path(destination) == self.config.output_path:
                self.config.output_path.write_bytes(sentinel)
            return real_operation(source, destination)

        with mock.patch.object(app.os, operation, side_effect=concurrent_create):
            with self.assertRaises(FileExistsError):
                app.run_pipeline(self.config)
        self.assertEqual(self.config.output_path.read_bytes(), sentinel)
        self.assertFalse(any(path.exists() for path in app.derive_delivery_paths(self.config).all_files()))

    def test_reviewed_input_cannot_be_replaced_by_a_historical_output(self) -> None:
        app.run_pipeline(self.config)
        paths = app.derive_delivery_paths(self.config)
        shutil.copyfile(self.config.output_path, paths.period_b_full)
        before = app._file_sha256(paths.period_b_full)
        with self.assertRaisesRegex(ValueError, "覆核輸入"):
            app.finalize_review_workbook(paths.period_b_full, delivery_dir=paths.root, overwrite=True)
        self.assertEqual(app._file_sha256(paths.period_b_full), before)

    def test_selected_reviewed_copy_is_locked_against_another_job(self) -> None:
        app.run_pipeline(self.config)
        reviewed = self.root / "reviewed_copy.xlsx"
        shutil.copyfile(self.config.output_path, reviewed)
        before = app._file_sha256(reviewed)
        entered, release = threading.Event(), threading.Event()
        failures = []

        def callback(payload):
            if payload.get("kind") == "preflight":
                entered.set()
                if not release.wait(15):
                    raise RuntimeError("review release timed out")

        def finalize():
            try:
                app.finalize_review_workbook(reviewed, delivery_dir=self.config.delivery_dir,
                                             overwrite=True, callback=callback)
            except BaseException as failure:
                failures.append(failure)

        thread = threading.Thread(target=finalize)
        thread.start()
        try:
            self.assertTrue(entered.wait(15))
            other = app.AppConfig(self.root, reviewed, "KLNG", 2026, 6,
                                  delivery_dir=self.root / "other_delivery", overwrite=True, workers=1)
            with self.assertRaisesRegex(RuntimeError, "另一個工作"):
                app.process_month(other)
        finally:
            release.set()
            thread.join(15)
        self.assertFalse(thread.is_alive())
        self.assertEqual(failures, [])
        self.assertEqual(app._file_sha256(reviewed), before)

    def test_gui_job_discovery_failure_does_not_disable_controls(self) -> None:
        # Exercise the real nested Tk callback without creating a desktop window.
        tree = ast.parse(Path(app.__file__).read_text(encoding="utf-8"))
        application = next(node for node in ast.walk(tree) if isinstance(node, ast.ClassDef) and node.name == "Application")
        start = next(node for node in application.body if isinstance(node, ast.FunctionDef) and node.name == "start")
        module = ast.Module(body=[start], type_ignores=[])
        messagebox = mock.Mock()
        namespace = {**app.__dict__, "messagebox": messagebox}
        exec(compile(module, str(app.__file__), "exec"), namespace)
        fake = SimpleNamespace(
            worker=None, build_config=lambda: self.config, cancel_event=threading.Event(),
            start_button=mock.Mock(), finalize_button=mock.Mock(), cancel_button=mock.Mock(),
            open_button=mock.Mock(), progress={}, status_var=mock.Mock(), progress_text_var=mock.Mock(),
            month_options={}, month_choice_var=mock.Mock(), append_log=mock.Mock(),
        )
        with mock.patch.dict(namespace, {"build_processing_job": mock.Mock(side_effect=app.SourceFileError("source changed before worker launch"))}):
            namespace["start"](fake)
        messagebox.showerror.assert_called_once()
        fake.start_button.configure.assert_not_called()
        fake.finalize_button.configure.assert_not_called()
        self.assertIsNone(fake.worker)

    def test_gui_start_overwrite_requires_existing_files_and_user_confirmation(self) -> None:
        for existing, answer in [(False, False), (True, False), (True, True)]:
            with self.subTest(existing=existing, answer=answer):
                output = self.root / f"start-{existing}-{answer}" / "analysis.xlsx"
                config = app.AppConfig(self.root, output, "KLNG", 2026, 6,
                                       overwrite=True, workers=1)
                sentinel = b"external-existing-output"
                if existing:
                    output.parent.mkdir(parents=True)
                    output.write_bytes(sentinel)
                start, fake, messagebox, _dialog, threads = self.gui_callback("start", config)
                messagebox.askyesno.return_value = answer
                start(fake)
                self.assertEqual(messagebox.askyesno.call_count, int(existing))
                if existing and not answer:
                    threads.assert_not_called()
                    self.assertEqual(output.read_bytes(), sentinel)
                    continue
                self.assertEqual(config.overwrite, existing and answer)
                if not existing:
                    output.parent.mkdir(parents=True, exist_ok=True)
                    output.write_bytes(sentinel)
                threads.call_args.kwargs["target"]()
                events = list(fake.events.queue)
                if not existing:
                    self.assertEqual(output.read_bytes(), sentinel)
                    self.assertEqual(events[-1]["kind"], "worker_error")
                else:
                    self.assertEqual(events[-1]["kind"], "worker_success")

    def test_gui_finalize_overwrite_requires_existing_files_and_user_confirmation(self) -> None:
        for existing, answer in [(False, False), (True, False), (True, True)]:
            with self.subTest(existing=existing, answer=answer):
                config = app.AppConfig(self.root, self.root / f"finalize-{existing}-{answer}" / "analysis.xlsx",
                                       "KLNG", 2026, 6, overwrite=True, workers=1)
                app.run_pipeline(config)
                paths = app.derive_delivery_paths(config)
                if not existing:
                    for path in paths.all_files():
                        path.unlink()
                finalize, fake, messagebox, dialog, threads = self.gui_callback("finalize_review", config)
                dialog.askopenfilename.return_value = str(config.output_path)
                messagebox.askyesno.return_value = answer
                before = app._file_sha256(config.output_path)
                finalize(fake)
                self.assertEqual(messagebox.askyesno.call_count, int(existing))
                if existing and not answer:
                    threads.assert_not_called()
                    continue
                sentinel = b"late-created-historical-output"
                if not existing:
                    paths.period_b_full.write_bytes(sentinel)
                threads.call_args.kwargs["target"]()
                events = list(fake.events.queue)
                if not existing:
                    self.assertEqual(paths.period_b_full.read_bytes(), sentinel)
                    self.assertEqual(events[-1]["kind"], "worker_error")
                else:
                    self.assertEqual(events[-1]["kind"], "worker_success")
                self.assertEqual(app._file_sha256(config.output_path), before)


    def test_all_historical_chart_titles_match_14pt_reference(self) -> None:
        import zipfile
        from xml.etree import ElementTree as ET

        results, _ = app.process_month(self.config)
        snapshot = app.build_default_decision_snapshot(self.config, results)
        paths = app.write_delivery_workbooks(self.config, results, snapshot)
        namespace = {
            "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        }
        for path in paths.all_files():
            with zipfile.ZipFile(path) as archive:
                charts = [name for name in archive.namelist()
                          if name.startswith("xl/charts/chart") and name.endswith(".xml")]
                self.assertEqual(len(charts), 2)
                for name in charts:
                    root = ET.fromstring(archive.read(name))
                    sizes = [element.get("sz") for element in
                             root.findall(".//c:title//a:defRPr", namespace)
                             if element.get("sz") is not None]
                    self.assertTrue(sizes, (path, name))
                    self.assertEqual(set(sizes), {"1400"}, (path, name))


if __name__ == "__main__":
    unittest.main()
