"""Run every catalogued port-month using the actual frozen executable.

Execution PASS is not workbook correctness PASS; use the full validator,
independent source census and native Excel QA afterwards. No day/file caps.
"""
from __future__ import annotations

import argparse
import ctypes
import hashlib
import json
import os
import re
import subprocess
import tempfile
import time
import zipfile
from datetime import datetime
from pathlib import Path

from ais_monthly_app import AppConfig, default_output_filename, derive_delivery_paths
from scripts.run_full_regression import discover_jobs, job_token, serializable_job
from scripts.verify_source_copy_manifest import atomic_json


def now() -> str:
    return datetime.now().astimezone().isoformat()


def guard_roots(source: Path, output: Path, executable: Path, *, resume: bool = False) -> None:
    source, output, executable = source.resolve(), output.resolve(), executable.resolve()
    if source == output or source in output.parents or output in source.parents:
        raise ValueError("Packaged regression output must not overlap the source tree")
    if output == executable.parent or output in executable.parents:
        raise ValueError("Packaged regression output must not contain the executable")
    if not executable.is_file() or executable.suffix.casefold() != ".exe":
        raise ValueError("Expected an existing Windows .exe")
    if output.exists() and (not output.is_dir() or (not resume and any(output.iterdir()))):
        raise ValueError("Fresh packaged regression requires an absent or empty output directory")


def executable_hash(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while block := handle.read(1024 * 1024):
            digest.update(block)
    return digest.hexdigest().upper()


def plan_hash(planned: list[dict[str, object]]) -> str:
    return hashlib.sha256(json.dumps(planned, sort_keys=True, ensure_ascii=False,
                                     separators=(",", ":")).encode("utf-8")).hexdigest().upper()


def windows_argv(command: str) -> list[str]:
    """Parse actual Windows argv, not embedded module names in a shell command."""
    if not command:
        return []
    count = ctypes.c_int()
    parser = ctypes.windll.shell32.CommandLineToArgvW
    parser.argtypes = [ctypes.c_wchar_p, ctypes.POINTER(ctypes.c_int)]
    parser.restype = ctypes.POINTER(ctypes.c_wchar_p)
    argv = parser(command, ctypes.byref(count))
    if not argv:
        raise ValueError("Cannot parse an inspected process command line")
    try:
        return [argv[index] for index in range(count.value)]
    finally:
        release = ctypes.windll.kernel32.LocalFree
        release.argtypes = [ctypes.c_void_p]
        release.restype = ctypes.c_void_p
        release(ctypes.cast(argv, ctypes.c_void_p))


def wrapper_targets(argv: list[str], output: Path) -> bool:
    # The normal module CLI has two positional roots immediately after -m.
    # Substring matching also matched the parent PowerShell diagnostic command.
    return (len(argv) >= 5 and argv[1:3] == ["-m", "scripts.run_packaged_full_regression"]
            and Path(argv[4]).resolve() == output.resolve())


def reject_live_writers(processes: list[dict[str, object]], output: Path, executable: Path) -> None:
    current = next((item for item in processes if item.get("ProcessId") == os.getpid()), None)
    if current is None:
        raise ValueError("Current process missing from Windows inspection; cannot prove ownership")
    current_args = windows_argv(str(current.get("CommandLine") or ""))
    parent_id = current.get("ParentProcessId")
    for process in processes:
        if process.get("ProcessId") == os.getpid():
            continue
        image = str(process.get("ExecutablePath") or "")
        name = str(process.get("Name") or "").casefold()
        if not image and name == executable.name.casefold():
            raise ValueError("Cannot inspect a process matching the producer executable name")
        if image and Path(image).resolve() == executable.resolve():
            raise ValueError(f"Producer/wrapper may still be live: PID {process.get('ProcessId')}")
        if name not in {"python.exe", "pythonw.exe"}:
            continue
        argv = windows_argv(str(process.get("CommandLine") or ""))
        if not wrapper_targets(argv, output):
            continue
        # Windows venv Python starts the base interpreter with identical CLI.
        # Exempt only our directly recorded parent, never another same-argv run.
        if (process.get("ProcessId") == parent_id and current_args == argv
                and wrapper_targets(current_args, output) and image):
            continue
        raise ValueError(f"Producer/wrapper may still be live: PID {process.get('ProcessId')}")


def assert_inactive(output: Path, executable: Path) -> None:
    """Fail closed; never kill writers or automatically remove producer locks."""
    if os.name != "nt":
        raise ValueError("Resume preparation requires Windows process inspection")
    query = ("[Console]::OutputEncoding=[System.Text.UTF8Encoding]::new(); "
             "Get-CimInstance Win32_Process | Select-Object ProcessId,ParentProcessId,Name,ExecutablePath,CommandLine | "
             "ConvertTo-Json -Compress")
    result = subprocess.run(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", query],
                            capture_output=True, text=True, encoding="utf-8", check=True)
    processes = json.loads(result.stdout)
    if isinstance(processes, dict):
        processes = [processes]
    reject_live_writers(processes, output, executable)
    lock_root = Path(tempfile.gettempdir()) / "AIS32_job_locks"
    for lock in lock_root.glob("*.lock"):
        metadata = json.loads(lock.read_text(encoding="utf-8"))
        resource = Path(str(metadata["resource"])).resolve()
        if resource == output.resolve() or output.resolve() in resource.parents:
            raise ValueError(f"Producer lock remains; inspect and archive manually: {lock}")


def contained(path: Path, root: Path) -> Path:
    resolved, boundary = path.resolve(), root.resolve()
    if resolved == boundary or boundary not in resolved.parents or path.is_symlink() or path.is_junction():
        raise ValueError(f"Noncanonical/out-of-scope artifact path: {path}")
    return resolved


def job_paths(job: dict[str, object], output: Path) -> tuple[Path, Path, list[Path], Path, Path]:
    token = job_token(job)
    if not re.fullmatch(r"[A-Z][A-Z0-9]*_\d{6}", token):
        raise ValueError(f"Invalid month token {token}")
    root = output / token
    modern = root / default_output_filename(str(job["port"]), int(job["year"]), int(job["month"]))
    delivery = root / "正式交付" / token
    config = AppConfig(Path(str(job["folder"])), modern, str(job["port"]), int(job["year"]),
                       int(job["month"]), delivery_dir=delivery)
    files = [modern, *derive_delivery_paths(config).all_files()]
    stem = f".{modern.stem}_{job['port']}_{job['year']}_{int(job['month']):02d}"
    return modern, delivery, files, root / (stem + "_cache"), root / (stem + "_rows")


def check_completed(record: dict[str, object], job: dict[str, object], output: Path) -> dict[str, str]:
    if (record.get("status") != "PASS" or record.get("exit_code") != 0 or
            record.get("observed_xlsx_count") != 6 or
            record.get("current_parser_logical_days") != job["tolerant_logical_days"] or
            any(record.get(key) != value for key, value in job.items())):
        raise ValueError(f"Invalid completed-month record: {job_token(job)}")
    modern, delivery, files, cache, spool = job_paths(job, output)
    if record.get("modern") != str(modern.resolve()) or record.get("delivery") != str(delivery.resolve()):
        raise ValueError("Noncanonical completed artifact paths")
    month_root = contained(modern.parent, output)
    for artifact in month_root.rglob("*"):
        contained(artifact, output)
    if set(month_root.rglob("*.xlsx")) != set(files):
        raise ValueError(f"Missing/extra final XLSX: {month_root}")
    fingerprints: dict[str, str] = {}
    for path in files:
        contained(path, output)
        if not path.is_file() or not zipfile.is_zipfile(path):
            raise ValueError(f"Invalid completed final XLSX: {path}")
        fingerprints[str(path.relative_to(output))] = executable_hash(path)
    expected_days = {str(day).replace("-", "") for day in job["logical_days"]}
    if (set(cache.iterdir()) != {cache / (day + ".json") for day in expected_days} or
            set(spool.iterdir()) != {spool / (day + ".bin") for day in expected_days}):
        raise ValueError(f"Missing/extra cache or spool days: {month_root}")
    sources = {str((Path(str(job["folder"])) / str(name)).resolve()) for name in job["source_files"]}
    observed_sources: set[str] = set()
    for day in sorted(expected_days):
        cache_path, spool_path = cache / f"{day}.json", spool / f"{day}.bin"
        contained(cache_path, output)
        contained(spool_path, output)
        payload = json.loads(cache_path.read_text(encoding="utf-8"))
        signature, result = payload["signature"], payload["result"]
        iso_day = f"{day[:4]}-{day[4:6]}-{day[6:]}"
        if (signature.get("date") != iso_day or result.get("day") != iso_day or
                any(signature.get(key) != job[key] for key in ("port", "year", "month")) or
                executable_hash(spool_path).casefold() != str(result.get("spool_sha256")).casefold()):
            raise ValueError(f"Cache/spool identity mismatch: {cache_path}")
        fragment_paths = {item["path"] for item in signature["sources"]}
        infos = {item["path"]: item for item in result["fragment_info"]}
        if (not fragment_paths or not fragment_paths <= sources or set(infos) != fragment_paths or
                set(result["source_fragments"]) != fragment_paths):
            raise ValueError(f"Cache source paths mismatch: {cache_path}")
        for item in signature["sources"]:
            source = Path(item["path"])
            if (str(source.resolve()) != item["path"] or source.stat().st_size != item["size"] or
                    source.stat().st_mtime_ns != item["modified_ns"] or
                    executable_hash(source).casefold() != str(infos[item["path"]]["sha256"]).casefold()):
                raise ValueError(f"Completed source content/signature changed: {source}")
        observed_sources.update(fragment_paths)
        for path in (cache_path, spool_path):
            fingerprints[str(path.relative_to(output))] = executable_hash(path)
    if observed_sources != sources:
        raise ValueError("Completed cache sources do not cover the saved full month plan")
    stored = record.get("resume_file_sha256")
    if stored is not None and stored != fingerprints:
        raise ValueError("Stored completed-month fingerprints changed")
    return fingerprints


def resume_checkpoint(source: Path, output: Path, executable: Path, planned: list[dict[str, object]],
                      *, prepare: bool) -> dict[str, object]:
    source, output, executable = source.resolve(), output.resolve(), executable.resolve()
    report_path = output / "run_report.json"
    contained(report_path, output)
    report = json.loads(report_path.read_text(encoding="utf-8"))
    if (report.get("scope") != "FULL" or report.get("source_root") != str(source) or
            report.get("planned_jobs") != planned or report.get("expected_job_count") != len(planned) or
            report.get("executable") != str(executable) or
            report.get("executable_sha256") != executable_hash(executable) or
            report.get("source_plan_sha256", plan_hash(planned)) != plan_hash(planned) or
            not report.get("started_at")):
        raise ValueError("Resume source plan/hash or executable identity mismatch")
    if not prepare and (report.get("status") != "PAUSED_READY" or not report.get("resume_prepared")):
        raise ValueError("Use explicit --prepare-resume after stopping old writers and manually archiving stale locks")
    assert_inactive(output, executable)
    by_token = {job_token(job): job for job in planned}
    records = report.get("jobs", [])
    tokens = [job_token(record) for record in records]
    if len(tokens) != len(set(tokens)) or any(token not in by_token for token in tokens):
        raise ValueError("Duplicate or unplanned resume records")
    completed: list[dict[str, object]] = []
    for record in records:
        if record.get("status") == "PASS":
            record["resume_file_sha256"] = check_completed(record, by_token[job_token(record)], output)
            completed.append(record)
        elif not prepare:
            raise ValueError("Prepared checkpoint contains a non-PASS completed record")
    complete_tokens = {job_token(record) for record in completed}
    pending = []
    for child in output.iterdir():
        if child == report_path:
            continue
        if not child.is_dir() or child.name not in by_token:
            raise ValueError(f"Unrecognized run-root entry; preserve and inspect manually: {child}")
        contained(child, output)
        if child.name not in complete_tokens:
            for artifact in child.rglob("*"):
                contained(artifact, output)
            pending.append(child)
    if not prepare:
        if pending:
            raise ValueError("Prepared checkpoint still contains partial-month artifacts")
        return report
    archive = output.with_name(output.name + ".resume_archive_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f"))
    guard_roots(source, archive, executable)
    if archive.parent.resolve() != output.parent.resolve() or archive == output or archive.exists():
        raise ValueError("Unsafe or existing recovery archive")
    # Validate every target first, then recheck inactivity immediately before
    # recoverable renames. Never delete caches, locks, reports or raw files.
    assert_inactive(output, executable)
    archive.mkdir()
    for child in pending:
        child.rename(archive / child.name)
    report_path.rename(archive / "run_report.json")
    report["jobs"] = completed
    report.pop("active_job", None)
    report.pop("error", None)
    report.pop("finished_at", None)
    report.update(status="PAUSED_READY", resume_prepared=True, source_plan_sha256=plan_hash(planned))
    report.setdefault("resume_events", []).append({"event": "PREPARED", "at": now(),
        "archive": str(archive), "restarted_whole_months": [child.name for child in pending],
        "retained_pass_months": sorted(complete_tokens),
        "fingerprint_boundary": "Legacy checkpoints first fingerprinted at preparation; formal content QA remains separate"})
    atomic_json(report_path, report)
    return report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source_root", type=Path)
    parser.add_argument("output_root", type=Path)
    parser.add_argument("--executable", required=True, type=Path)
    parser.add_argument("--workers", type=int, default=3)
    parser.add_argument("--pause-flag", type=Path,
                        help="User quota stop marker; never start another month when present")
    continuation = parser.add_mutually_exclusive_group()
    continuation.add_argument("--prepare-resume", action="store_true")
    continuation.add_argument("--resume", action="store_true")
    args = parser.parse_args(argv)
    if not 1 <= args.workers <= 8:
        parser.error("workers must be between 1 and 8")
    args.source_root, args.output_root, args.executable = (path.resolve() for path in
        (args.source_root, args.output_root, args.executable))
    if args.pause_flag is not None:
        args.pause_flag = args.pause_flag.resolve()
        for protected in (args.source_root, args.output_root, args.executable.parent):
            if args.pause_flag == protected or protected in args.pause_flag.parents:
                raise ValueError("Pause marker must be outside source/output/executable trees")
    guard_roots(args.source_root, args.output_root, args.executable, resume=args.resume or args.prepare_resume)
    planned = [serializable_job(job) for job in discover_jobs(args.source_root)]
    if not planned:
        raise ValueError("No applicable source jobs; cannot pass")
    if args.prepare_resume:
        report = resume_checkpoint(args.source_root, args.output_root, args.executable, planned, prepare=True)
        print(json.dumps({"status": report["status"], "completed": len(report["jobs"]),
                          "archive": report["resume_events"][-1]["archive"]}, ensure_ascii=False))
        return 0
    args.output_root.mkdir(parents=True, exist_ok=True)
    report_path = args.output_root / "run_report.json"
    records: list[dict[str, object]] = []
    initial_hash = executable_hash(args.executable)
    report: dict[str, object] = {
        "source_root": str(args.source_root.resolve()), "started_at": now(),
        "status": "IN_PROGRESS", "scope": "FULL", "planned_jobs": planned,
        "expected_job_count": len(planned), "jobs": records,
        "executable": str(args.executable.resolve()), "executable_sha256": initial_hash,
        "pass_boundary": "Execution only; workbook/source/Excel/visual validation is separate",
        "source_plan_sha256": plan_hash(planned),
    }
    if args.resume:
        report = resume_checkpoint(args.source_root, args.output_root, args.executable, planned, prepare=False)
        records = report["jobs"]
        report.update(status="IN_PROGRESS", resume_prepared=False)
        report.setdefault("resume_events", []).append({"event": "RESUMED", "at": now()})
    retained = {job_token(record) for record in records}
    atomic_json(report_path, report)
    def user_paused(stage: str) -> bool:
        if args.pause_flag is None or not args.pause_flag.exists():
            return False
        report.update(status="PAUSED_BY_USER_QUOTA", paused_at=now(),
                      pause_flag=str(args.pause_flag), pause_stage=stage,
                      pause_boundary="Partial artifacts retained; interrupted month is not data FAIL or completed")
        atomic_json(report_path, report)
        print(json.dumps({"status": report["status"], "stage": stage,
                          "completed_months": len(records)}, ensure_ascii=False), flush=True)
        return True
    try:
        for position, job in enumerate(planned, 1):
            if user_paused("BEFORE_NEXT_MONTH"):
                return 2
            token = job_token(job)
            if token in retained:
                print(f"[{position}/{len(planned)}] retained validated completed {token}", flush=True)
                continue
            report["active_job"] = token
            atomic_json(report_path, report)
            job_root = args.output_root / token
            job_root.mkdir(parents=True, exist_ok=True)
            modern = job_root / default_output_filename(str(job["port"]), int(job["year"]), int(job["month"]))
            delivery = job_root / "正式交付" / token
            command = [str(args.executable.resolve()), "--input", str(job["folder"]),
                       "--output", str(modern), "--delivery-dir", str(delivery),
                       "--port", str(job["port"]), "--year", str(job["year"]),
                       "--month", str(job["month"]), "--workers", str(args.workers), "--overwrite"]
            print(f"[{position}/{len(planned)}] packaged {token}", flush=True)
            started = time.perf_counter()
            with (job_root / "exe.stdout.log").open("wb") as stdout, (job_root / "exe.stderr.log").open("wb") as stderr:
                result = subprocess.run(command, stdout=stdout, stderr=stderr, check=False)
            # Do not classify quota-stopped child as data FAIL or race into the
            # next month while the owning process tree is being stopped.
            if user_paused("AFTER_ACTIVE_CHILD"):
                return 2
            actual = list(job_root.rglob("*.xlsx"))
            cache_days = [path for path in job_root.rglob("*.json")
                          if path.parent.name.endswith("_cache") and re.fullmatch(r"\d{8}\.json", path.name)]
            record = {**job, "modern": str(modern), "delivery": str(delivery),
                      "current_parser_logical_days": len(cache_days),
                      "detected_day_count_boundary": "Actual EXE-created per-day cache count; workbook dates/content validated separately",
                      "exit_code": result.returncode, "observed_xlsx_count": len(actual),
                      "status": "PASS" if result.returncode == 0 and len(actual) == 6
                                and len(cache_days) == job["tolerant_logical_days"] else "FAIL",
                      "elapsed_seconds": round(time.perf_counter() - started, 3), "finished_at": now()}
            if record["status"] == "FAIL":
                record["error"] = (f"EXE exit={result.returncode}; expected 6 final XLSX, observed {len(actual)}; "
                                   f"expected {job['tolerant_logical_days']} cache days, observed {len(cache_days)}")
            records.append(record)
            atomic_json(report_path, report)
            print(json.dumps(record, ensure_ascii=False), flush=True)
        if executable_hash(args.executable) != initial_hash:
            raise ValueError("Executable changed during regression")
    except BaseException as error:
        report.update(status="INTERRUPTED" if isinstance(error, (KeyboardInterrupt, SystemExit)) else "FAILED",
                      error=f"{type(error).__name__}: {error}", finished_at=now())
        atomic_json(report_path, report)
        raise
    report.pop("active_job", None)
    report.update(status="COMPLETE", finished_at=now())
    atomic_json(report_path, report)
    return 1 if any(item["status"] != "PASS" for item in records) else 0


if __name__ == "__main__":
    raise SystemExit(main())
