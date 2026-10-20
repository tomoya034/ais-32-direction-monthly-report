"""Read-only per-completed-month serialization checkpoint; partial scope never global PASS."""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from scripts.audit_default_workbook_content import audit_job
from scripts.independent_source_audit import guard_output, save_report


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("root", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--token", action="append", default=[])
    args = parser.parse_args(argv)
    run_file = args.root / "run_report.json"
    run = json.loads(run_file.read_text(encoding="utf-8"))
    protected = [args.root, Path(run["source_root"])]
    guard_output(args.output, protected, [run_file])
    if args.output.exists():
        raise ValueError("existing evidence must not be blindly overwritten; choose a fresh checkpoint filename")
    planned = {job["token"]: job for job in run["planned_jobs"]}
    completed = [job for job in run["jobs"] if job.get("status") == "PASS"
                 and job.get("exit_code") == 0 and job.get("observed_xlsx_count") == 6]
    if len({job["token"] for job in completed}) != len(completed):
        raise ValueError("duplicate completed month cannot pass")
    if args.token:
        wanted = set(args.token)
        if wanted - planned.keys() or wanted - {job["token"] for job in completed}:
            raise ValueError("requested month is unplanned or incomplete")
        completed = [job for job in completed if job["token"] in wanted]
    if not completed:
        raise ValueError("no completed months to verify")
    report = {"schema": "AIS_COMPLETED_DEFAULT_MONTH_CHECKPOINT_1", "root": str(args.root.resolve()),
              "source_root": run["source_root"], "planned_tokens": sorted(planned),
              "selected_tokens": [job["token"] for job in completed], "jobs": [], "pass": False,
              "boundary": "Serialization only; raw source-key and independent selector/candidate, native/visual, "
                          "immutable source and full repeated-run gates remain separate. Partial months never included as PASS."}
    save_report(args.output, report, protected, [run_file])
    for job in completed:
        try:
            result = audit_job(job)
        except Exception as error:
            result = {"token": job["token"], "pass": False, "issues": [f"{type(error).__name__}: {error}"]}
        report["jobs"].append(result)
        report["pass"] = (run["status"] == "COMPLETE" and set(report["selected_tokens"]) == set(planned)
                          and len(report["jobs"]) == len(planned) and all(item["pass"] for item in report["jobs"]))
        save_report(args.output, report, protected, [run_file])
        print(json.dumps({"token": result["token"], "pass": result["pass"], "issues": result["issues"]}, ensure_ascii=False), flush=True)
    print(json.dumps({"global_pass": report["pass"], "months_checked": len(report["jobs"]),
                      "remaining_tokens": sorted(set(planned) - set(report["selected_tokens"]))}, ensure_ascii=False), flush=True)
    return 0 if all(item["pass"] for item in report["jobs"]) else 1


if __name__ == "__main__":
    raise SystemExit(main())
