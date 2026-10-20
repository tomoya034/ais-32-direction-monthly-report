"""Revalidate same-root sealed selector checkpoints; missing months remain UNVERIFIED."""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from scripts.independent_selector_audit import retained_results
from scripts.independent_source_audit import guard_output, plan_tasks, save_report


def collect(run_root, report_paths):
    run = json.loads((run_root / "run_report.json").read_text(encoding="utf-8"))
    planned = {job["token"]: job for job in run["planned_jobs"]}
    if len(planned) != len(run["planned_jobs"]):
        raise ValueError("duplicate planned port-month")
    complete = {job["token"] for job in run["jobs"] if job.get("status") == "PASS"
                and job.get("exit_code") == 0 and job.get("observed_xlsx_count") == 6}
    summaries, seen, keys = [], set(), set()
    for path in report_paths:
        report = json.loads(path.read_text(encoding="utf-8"))
        token = Path(report["run_root"]).name
        if token in seen or token not in planned or token not in complete:
            raise ValueError("unplanned, duplicate, or incomplete month checkpoint")
        seen.add(token)
        source = Path(planned[token]["folder"])
        month_root = run_root / token
        tasks = plan_tasks(source, month_root)
        valid = retained_results(report, tasks, source, month_root)
        day_keys = {(entry["port"], entry["day"]) for entry in valid}
        if (report.get("pass") is not True or len(valid) != len(tasks) or keys & day_keys
                or len(day_keys) != len(valid)):
            raise ValueError("incomplete selector month or duplicate logical day")
        keys.update(day_keys)
        summaries.append({"token": token, "report": str(path.resolve()), "pass": True,
                          "days": len(valid), "rows": sum(entry["rows_scanned"] for entry in valid),
                          "numeric": sum(entry["numeric_fields_checked"] for entry in valid),
                          "ids": sum(entry["identity_fields_checked"] for entry in valid),
                          "candidates": sum(entry["candidate_rows_checked"] for entry in valid)})
    totals = {key: sum(item[key] for item in summaries) for key in ("days", "rows", "numeric", "ids", "candidates")}
    full = run.get("status") == "COMPLETE" and seen == set(planned) and complete == set(planned)
    if full:
        full_tasks = plan_tasks(Path(run["source_root"]), run_root)
        if keys != {(task["port"], task["day"]) for task in full_tasks}:
            raise ValueError("global source/cache/selector logical-day coverage differs")
    return {"schema": "AIS_SELECTOR_MONTH_COLLECTOR_1", "root": str(run_root.resolve()),
            "jobs": summaries, "totals": totals, "pass": full,
            "remaining_tokens": sorted(set(planned) - seen),
            "post_seal_stamps_cache_rechecked": True,
            "boundary": "Selector/candidate automatic-rule evidence only. Same-root independent RAW SOURCE-KEY "
                        "census, final workbook/bindings, native/visual, source SHA and repeat gates required separately."}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("root", type=Path)
    parser.add_argument("reports", nargs="+", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args(argv)
    run_file = args.root / "run_report.json"
    run = json.loads(run_file.read_text(encoding="utf-8"))
    protected = [args.root, Path(run["source_root"])]
    guard_output(args.output, protected, [run_file, *args.reports])
    if args.output.exists():
        raise ValueError("existing collector evidence must not be blindly replaced")
    report = collect(args.root, args.reports)
    save_report(args.output, report, protected, [run_file, *args.reports])
    print(json.dumps(report, ensure_ascii=False))
    return 0  # Partial coverage is explicitly represented by pass=false, not omitted.


if __name__ == "__main__":
    raise SystemExit(main())
