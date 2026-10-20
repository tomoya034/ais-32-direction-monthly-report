from __future__ import annotations

import argparse
import contextlib
import datetime as dt
import hashlib
import json
import os
import re
import stat
import sys
import tempfile
import time
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator


CHECKPOINT_VERSION = 2


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def signature(value: os.stat_result) -> dict[str, int]:
    # Windows st_ctime is deprecated and Path.stat/os.fstat can disagree on
    # its meaning. Python documents st_birthtime_ns as the creation timestamp.
    # Content modification and identity remain guarded by mtime/size/dev/ino.
    time_field = ("st_birthtime_ns" if os.name == "nt" and hasattr(value, "st_birthtime_ns")
                  else "st_ctime_ns")
    return {name: int(getattr(value, name)) for name in
            ("st_size", "st_mtime_ns", time_field, "st_dev", "st_ino")}


class FileChangedError(OSError):
    pass


def hash_file_stable(path: Path) -> tuple[str, int, dict[str, int]]:
    """Hash one regular file and reject observable changes during the read."""
    before_path = path.stat()
    digest = hashlib.sha256()
    byte_count = 0
    with path.open("rb") as handle:
        before_handle = os.fstat(handle.fileno())
        if not stat.S_ISREG(before_handle.st_mode):
            raise OSError("not a regular file")
        if signature(before_path) != signature(before_handle):
            raise FileChangedError("file changed between stat and open")
        while chunk := handle.read(1024 * 1024):
            byte_count += len(chunk)
            if byte_count > before_handle.st_size:
                raise FileChangedError("file grew during hashing")
            digest.update(chunk)
        after_handle = os.fstat(handle.fileno())
        after_path = path.stat()
    expected_signature = signature(before_handle)
    if (signature(after_handle) != expected_signature
            or signature(after_path) != expected_signature
            or byte_count != before_handle.st_size):
        raise FileChangedError("file changed during hashing")
    return digest.hexdigest().upper(), byte_count, expected_signature


def atomic_json(path: Path, value: object) -> None:
    """Keep the previous complete JSON if writing or replacement is interrupted."""
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", dir=path.parent,
            prefix=path.name + ".", suffix=".tmp", delete=False,
        ) as handle:
            temporary = Path(handle.name)
            json.dump(value, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


@contextlib.contextmanager
def exclusive_lock(checkpoint: Path) -> Iterator[None]:
    """An OS lock is released after a crash; the lock file itself may remain."""
    lock_path = checkpoint.with_name(checkpoint.name + ".lock")
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    with lock_path.open("a+b") as handle:
        if handle.seek(0, os.SEEK_END) == 0:
            handle.write(b"\0")
            handle.flush()
        handle.seek(0)
        locked = False
        try:
            if os.name == "nt":
                import msvcrt
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            locked = True
            yield
        finally:
            if locked:
                handle.seek(0)
                if os.name == "nt":
                    import msvcrt
                    msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(handle.fileno(), fcntl.LOCK_UN)


@dataclass(frozen=True)
class Task:
    scope: str
    relative: str
    path: Path
    expected_size: int
    expected_hash: str

    @property
    def key(self) -> str:
        return f"{self.scope}::{self.relative}"


def build_tasks(manifest: dict, scopes: tuple[str, ...]) -> list[Task]:
    entries = manifest["files"]
    if not isinstance(entries, list) or not entries:
        raise ValueError("manifest files must be a nonempty list")
    seen: set[str] = set()
    tasks: list[Task] = []
    for entry in entries:
        relative = Path(entry["relative_path"])
        if relative.is_absolute() or relative.drive or ".." in relative.parts:
            raise ValueError(f"unsafe relative path: {relative}")
        normalized = str(relative)
        if relative.suffix.lower() != ".xlsx" or normalized.casefold() in seen:
            raise ValueError(f"non-XLSX or duplicate relative path: {relative}")
        seen.add(normalized.casefold())
        if len(scopes) == 2 and (
            int(entry["source_size"]) != int(entry["copy_size"])
            or str(entry["source_sha256"]).upper() != str(entry["copy_sha256"]).upper()
        ):
            raise ValueError(f"source/copy expected content differs in baseline: {relative}")
        for scope in scopes:
            size = int(entry[f"{scope}_size"])
            expected_hash = str(entry[f"{scope}_sha256"]).upper()
            if size < 0 or not re.fullmatch(r"[0-9A-F]{64}", expected_hash):
                raise ValueError(f"invalid size/hash: {scope}::{relative}")
            tasks.append(Task(scope, normalized, Path(manifest[scope]) / relative,
                              size, expected_hash))
    return tasks


def task_fingerprint(tasks: list[Task]) -> str:
    value = [(t.scope, t.relative, str(t.path.resolve()), t.expected_size,
              t.expected_hash) for t in tasks]
    return hashlib.sha256(json.dumps(value, ensure_ascii=False,
                                    separators=(",", ":")).encode()).hexdigest().upper()


def inventory(root: Path, expected: set[str]) -> dict:
    actual: set[str] = set()
    errors: list[str] = []
    if not root.is_dir():
        errors.append("root is not an accessible directory")
    else:
        for directory, _, names in os.walk(root, onerror=lambda e: errors.append(str(e))):
            for name in names:
                if Path(name).suffix.lower() == ".xlsx":
                    actual.add(str((Path(directory) / name).relative_to(root)))
    return {"actual_file_count": len(actual), "missing": sorted(expected - actual),
            "extra": sorted(actual - expected), "errors": errors}


def verify(task: Task, retries: int) -> dict:
    result = {"scope": task.scope, "relative_path": task.relative,
              "expected_size": task.expected_size, "expected_sha256": task.expected_hash,
              "hash_attempt_count": 0}
    for attempt in range(1, retries + 1):
        try:
            result["hash_attempt_count"] += 1
            actual_hash, actual_size, actual_signature = hash_file_stable(task.path)
            result.update(actual_size=actual_size, actual_sha256=actual_hash,
                          signature=actual_signature, verified_at=utc_now(), attempts=attempt)
            result.pop("error", None)
            result["status"] = ("match" if actual_size == task.expected_size
                                and actual_hash == task.expected_hash else "mismatch")
            return result
        except OSError as error:
            result.update(status="error", error=f"{type(error).__name__}: {error}")
            result.setdefault("retry_errors", []).append(result["error"])
            if attempt < retries:
                time.sleep(min(5 * attempt, 20))
    result["attempts"] = retries
    return result


def reusable(task: Task, record: object) -> bool:
    if not isinstance(record, dict) or record.get("status") != "match":
        return False
    if (record.get("scope") != task.scope or record.get("relative_path") != task.relative
            or record.get("expected_size") != task.expected_size
            or record.get("expected_sha256") != task.expected_hash
            or record.get("actual_size") != task.expected_size
            or record.get("actual_sha256") != task.expected_hash
            or not record.get("verified_at")):
        return False
    try:
        return signature(task.path.stat()) == record.get("signature")
    except OSError:
        return False


def validate_artifact_locations(manifest: dict, manifest_path: Path,
                                output: Path, checkpoint: Path, tasks: list[Task]) -> None:
    """Reports, progress, locks and their temporary writers must stay outside inputs."""
    roots = [Path(manifest[scope]).resolve() for scope in ("source", "copy") if scope in manifest]
    protected_files = {manifest_path.resolve(), *(task.path.resolve() for task in tasks)}
    artifacts = (output, checkpoint, checkpoint.with_name(checkpoint.name + ".lock"))
    resolved = [path.resolve() for path in artifacts]
    if len(set(resolved)) != len(resolved):
        raise ValueError("report/checkpoint/lock paths must be distinct")
    for path, target in zip(artifacts, resolved):
        # Check the parent too: an existing final symlink may point outside a raw
        # root while atomic_json would still create its temporary file inside it.
        parent = path.parent.resolve()
        if target in protected_files or any(
            target == root or root in target.parents or parent == root or root in parent.parents
            for root in roots
        ):
            raise ValueError(f"audit artifact must stay outside source/copy roots and input files: {path}")


def run(arguments: argparse.Namespace) -> int:
    manifest_bytes = arguments.manifest.read_bytes()
    manifest = json.loads(manifest_bytes.decode("utf-8"))
    manifest_hash = hashlib.sha256(manifest_bytes).hexdigest().upper()
    scopes = ("source", "copy") if arguments.scope == "both" else (arguments.scope,)
    tasks = build_tasks(manifest, scopes)
    fingerprint = task_fingerprint(tasks)
    checkpoint_path = arguments.checkpoint or arguments.output.with_suffix(".progress.json")
    validate_artifact_locations(manifest, arguments.manifest, arguments.output, checkpoint_path, tasks)
    report = {"run_id": str(uuid.uuid4()), "started_at": utc_now(),
              "manifest": str(arguments.manifest.resolve()), "manifest_sha256": manifest_hash,
              "task_fingerprint": fingerprint, "scope": arguments.scope,
              "expected_file_count": len(manifest["files"]), "task_count": len(tasks),
              "fresh_requested": arguments.fresh, "roots": {}, "events": [],
              "resume_semantics": "Resumed results are prior hashes with unchanged file metadata; "
              "they are not fresh content hashes. Use --fresh for a full new content check.",
              "snapshot_semantics": "Per-file hashes and final metadata/inventory checks are "
              "observations, not an atomic filesystem-wide NAS snapshot."}
    events = report["events"]
    completed: dict[str, dict] = {}
    with exclusive_lock(checkpoint_path):
        if checkpoint_path.is_file() and not arguments.fresh:
            try:
                saved = json.loads(checkpoint_path.read_text(encoding="utf-8"))
                if (saved.get("version") == CHECKPOINT_VERSION
                        and saved.get("manifest_sha256") == manifest_hash
                        and saved.get("task_fingerprint") == fingerprint
                        and saved.get("scope") == arguments.scope
                        and isinstance(saved.get("completed"), dict)):
                    for task in tasks:
                        record = saved["completed"].get(task.key)
                        if reusable(task, record):
                            completed[task.key] = record
                else:
                    events.append({"type": "checkpoint_incompatible", "action": "rehash"})
            except (ValueError, AttributeError, OSError) as error:
                events.append({"type": "checkpoint_unreadable", "action": "rehash",
                               "error": f"{type(error).__name__}: {error}"})
        resumed_count = len(completed)
        expected = {t.relative for t in tasks}
        for scope in scopes:
            root = Path(manifest[scope])
            report["roots"][scope] = {"path": str(root),
                                       "start_inventory": inventory(root, expected)}

        def checkpoint() -> None:
            atomic_json(checkpoint_path, {"version": CHECKPOINT_VERSION,
                "manifest": str(arguments.manifest.resolve()), "manifest_sha256": manifest_hash,
                "scope": arguments.scope, "task_fingerprint": fingerprint,
                "updated_at": utc_now(), "completed": completed})

        pending = [t for t in tasks if t.key not in completed]
        current_results: list[dict] = []
        with ThreadPoolExecutor(max_workers=arguments.workers) as executor:
            futures = {executor.submit(verify, t, arguments.retries): t for t in pending}
            for ordinal, future in enumerate(as_completed(futures), start=1):
                task = futures[future]
                result = future.result()
                current_results.append(result)
                completed[task.key] = result
                if ordinal % 10 == 0 or ordinal == len(pending):
                    checkpoint()
                    print(f"checked {len(completed)}/{len(tasks)} "
                          f"(fresh {ordinal}, resumed {resumed_count})", flush=True)

        inventory_ok = True
        for scope in scopes:
            root_record = report["roots"][scope]
            ending = inventory(Path(manifest[scope]), expected)
            root_record.update(ending)
            root_record["end_inventory"] = ending
            root_record["inventory_changed"] = root_record["start_inventory"] != ending
            inventory_ok &= all(not scan["missing"] and not scan["extra"] and not scan["errors"]
                                for scan in (root_record["start_inventory"], ending))
            inventory_ok &= not root_record["inventory_changed"]
        try:
            manifest_unchanged = hashlib.sha256(arguments.manifest.read_bytes()).hexdigest().upper() == manifest_hash
        except OSError as error:
            manifest_unchanged = False
            events.append({"type": "manifest_unreadable_at_end", "error": str(error)})
        if not manifest_unchanged:
            events.append({"type": "manifest_changed_during_run"})
        # These must follow the ending inventory/manifest reads: those scans can
        # take time, during which an already-hashed file may change again.
        for task in tasks:
            record = completed[task.key]
            if record.get("status") == "match" and not reusable(task, record):
                record = dict(record)
                record.update(status="error", error="file changed or became inaccessible after verification")
                completed[task.key] = record
        mismatches = [completed[t.key] for t in tasks if completed[t.key].get("status") != "match"]
        # An unreadable/short/unstable file has an attempt, not a completed hash.
        hashed_now = sum("actual_sha256" in r for r in current_results)
        report.update(completed_at=utc_now(), hashed_file_count=hashed_now,
                      hashed_this_run=hashed_now, attempted_this_run=len(pending),
                      hash_attempt_count=sum(r["hash_attempt_count"] for r in current_results),
                      resumed_success_count=resumed_count, checked_file_count=len(completed),
                      match_count=len(tasks) - len(mismatches), mismatches=mismatches,
                      manifest_unchanged=manifest_unchanged,
                      **{"pass": inventory_ok and manifest_unchanged and not mismatches})
        checkpoint()
        atomic_json(arguments.output, report)
    print(json.dumps({key: report[key] for key in ("pass", "task_count", "hashed_this_run",
          "resumed_success_count", "match_count")}, ensure_ascii=False, indent=2))
    return 0 if report["pass"] else 1


def positive(value: str) -> int:
    number = int(value)
    if number < 1:
        raise argparse.ArgumentTypeError("must be at least 1")
    return number


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Verify immutable AIS source/copy; metadata-based "
                                               "resume is distinct from a fresh content hash.")
    parser.add_argument("manifest", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--scope", choices=("source", "copy", "both"), default="both")
    parser.add_argument("--workers", type=positive, default=2)
    parser.add_argument("--retries", type=positive, default=5)
    parser.add_argument("--checkpoint", type=Path)
    parser.add_argument("--fresh", action="store_true", help="Do not reuse any checkpoint hashes.")
    arguments = parser.parse_args(argv)
    try:
        return run(arguments)
    except (OSError, ValueError, KeyError, TypeError) as error:
        print(f"Verification did not complete: {type(error).__name__}: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
