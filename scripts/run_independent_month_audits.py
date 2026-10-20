"""Collect completed independent month audits; never launch or overwrite a month audit.

The producer may still be writing its atomic daily checkpoint. Missing/partial
months remain IN_PROGRESS. Final aggregation rechecks every sealed input stamp,
current cache counter and record census with the independent audit module.
"""
from __future__ import annotations

import argparse
from datetime import datetime, timezone
import json
from pathlib import Path
import time

from scripts import independent_source_audit as independent


def source_months(source_root: Path) -> list[dict]:
    if not source_root.is_dir():
        raise ValueError(f"source root does not exist: {source_root}")
    grouped: dict[str, dict] = {}
    for path in sorted(source_root.rglob("*.xlsx")):
        port, day = independent.independent_identity(path.name)
        token = f"{port}_{day:%Y%m}"
        month = grouped.setdefault(token, {"token": token, "folder": str(path.parent.resolve()),
                                          "files": [], "days": set()})
        if month["folder"] != str(path.parent.resolve()):
            raise ValueError(f"month is spread across multiple folders: {token}")
        month["files"].append(str(path.resolve()))
        month["days"].add(day.isoformat())
    if not grouped:
        raise ValueError("no source months; empty scope cannot pass")
    return [{**month, "days": sorted(month["days"])} for _, month in sorted(grouped.items())]


def collect(source_root: Path, run_root: Path, report_directory: Path,
            prefix: str = "independent_raw_v155_", expected: dict | None = None) -> dict:
    source_root, run_root = source_root.resolve(), run_root.resolve()
    months = source_months(source_root)
    counts = {"source_files": sum(len(month["files"]) for month in months),
              "logical_days": sum(len(month["days"]) for month in months), "months": len(months)}
    issues: list[str] = []
    pending: list[str] = []
    results: list[dict] = []
    report = {"checkpoint_schema": independent.CHECKPOINT_SCHEMA,
              "source_root": str(source_root), "run_root": str(run_root),
              "expected_source_files": counts["source_files"],
              "expected_logical_days": counts["logical_days"], "expected_months": counts["months"],
              "planned_months": months, "results": results, "issues": issues, "pending_months": pending,
              "scope": "all sealed independent raw/cache/spool month audits; selector/manual decisions/charts excluded",
              "resume_boundary": "Sealed checksums and current input stat stamps/census; not signed authentication "
              "or an atomic NAS snapshot. Separate immutable source SHA audit remains required."}
    for name, value in (expected or {}).items():
        if value is not None and counts[name] != value:
            issues.append(f"source plan {name} {counts[name]} != required {value}")
    for month in months:
        token = month["token"]
        path = report_directory / f"{prefix}{token}.json"
        if not path.is_file():
            pending.append(token)
            continue
        try:
            previous = json.loads(path.read_text(encoding="utf-8"))
            entries = previous.get("results", [])
            if any(isinstance(entry, dict) and entry.get("pass") is False for entry in entries):
                raise ValueError("month retains failed day results; investigate/retry before collecting")
            if not previous.get("pass"):
                pending.append(token)
                continue
            if (previous.get("expected_source_files") != len(month["files"])
                    or previous.get("expected_logical_days") != len(month["days"])
                    or previous.get("completed_days") != len(month["days"])):
                raise ValueError("month completion/source counts differ from independent filename plan")
            tasks = independent.plan_tasks(Path(month["folder"]), run_root / token)
            retained = independent.resume_results(previous, tasks, Path(month["folder"]), run_root / token)
            if len(retained) != len(month["days"]):
                raise ValueError("month does not contain every sealed PASS day")
            results.extend(retained)
        except Exception as error:
            issues.append(f"{token}: {type(error).__name__}: {error}")
    if not issues and not pending:
        try:
            tasks = independent.plan_tasks(source_root, run_root)
            results[:] = independent.resume_results(report, tasks, source_root, run_root)
            if (len(tasks) != counts["logical_days"] or len(results) != counts["logical_days"]
                    or sum(len(task["paths"]) for task in tasks) != counts["source_files"]
                    or len({f"{task['port']}_{str(task['day'])[:7].replace('-', '')}" for task in tasks})
                    != counts["months"]):
                raise ValueError("global tasks/results are not the complete source filename scope")
        except Exception as error:
            issues.append(f"global revalidation: {type(error).__name__}: {error}")
    results.sort(key=lambda item: (str(item["port"]), str(item["day"])))
    report.update(completed_days=len(results), checked_at=datetime.now(timezone.utc).isoformat(),
                  status="FAILED" if issues else "IN_PROGRESS" if pending else "COMPLETE",
                  **{"pass": not issues and not pending and len(results) == counts["logical_days"]})
    return report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source_root", type=Path)
    parser.add_argument("run_root", type=Path)
    parser.add_argument("report_directory", type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--prefix", default="independent_raw_v155_")
    parser.add_argument("--expect-files", type=int)
    parser.add_argument("--expect-days", type=int)
    parser.add_argument("--expect-months", type=int)
    parser.add_argument("--wait", action="store_true", help="Wait for month producers; never start them.")
    parser.add_argument("--poll-seconds", type=float, default=30)
    parser.add_argument("--cancel-file", type=Path)
    args = parser.parse_args(argv)
    if args.poll_seconds <= 0 or any(value is not None and value < 1 for value in
                                  (args.expect_files, args.expect_days, args.expect_months)):
        parser.error("poll interval and required counts must be positive")
    months = source_months(args.source_root)
    protected = [Path(path) for month in months for path in month["files"]]
    protected.extend(args.report_directory / f"{args.prefix}{month['token']}.json" for month in months)
    protected.extend(path for path in args.run_root.rglob("*") if path.is_file())
    independent.guard_output(args.output, [args.source_root, args.run_root], protected)
    while True:
        if args.cancel_file is not None and args.cancel_file.exists():
            # No month producer is owned by this collector. Its daily checkpoints
            # remain untouched; the operator controls any separately started PID.
            report = {"status": "PAUSED", "pass": False, "cancel_file": str(args.cancel_file),
                      "source_root": str(args.source_root.resolve()), "run_root": str(args.run_root.resolve())}
        else:
            report = collect(args.source_root, args.run_root, args.report_directory, args.prefix,
                             {"source_files": args.expect_files, "logical_days": args.expect_days,
                              "months": args.expect_months})
        independent.save_report(args.output, report, [args.source_root, args.run_root], protected)
        print(json.dumps({key: report.get(key) for key in ("status", "pass", "completed_days", "pending_months", "issues")},
                         ensure_ascii=False), flush=True)
        if report["pass"]:
            return 0
        if report["status"] == "FAILED":
            return 1
        if not args.wait or report["status"] == "PAUSED":
            return 2
        # Check cancellation in short intervals even when waiting minutes.
        deadline = time.monotonic() + args.poll_seconds
        while time.monotonic() < deadline:
            if args.cancel_file is not None and args.cancel_file.exists():
                break
            time.sleep(max(0.0, min(1.0, deadline - time.monotonic())))


if __name__ == "__main__":
    raise SystemExit(main())
