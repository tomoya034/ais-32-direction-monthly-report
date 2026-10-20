from __future__ import annotations

import argparse
import copy
import hashlib
import json
import time
import traceback
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from ais_monthly_app import (
    AppConfig,
    SOURCE_LIKE_FILENAME_RE,
    default_output_filename,
    detect_source_catalog,
    run_pipeline,
)
from scripts.audit_source_catalog import tolerant_identity
from scripts.verify_source_copy_manifest import atomic_json
from scripts.independent_source_audit import guard_output


def job_token(job: dict[str, object]) -> str:
    return f"{job['port']}_{int(job['year'])}{int(job['month']):02d}"


def serializable_job(job: dict[str, object]) -> dict[str, object]:
    return copy.deepcopy({**{key: (str(value) if isinstance(value, Path) else value)
                           for key, value in job.items()}, "token": job_token(job)})


def discover_jobs(source_root: Path) -> list[dict[str, object]]:
    if not source_root.is_dir():
        raise ValueError(f"來源根目錄不存在：{source_root}")
    grouped: defaultdict[Path, list[tuple[Path, tuple[str, str, str | None]]]] = defaultdict(list)
    for path in source_root.rglob("*.xlsx"):
        if path.name.startswith("~$"):
            continue
        identity = tolerant_identity(path.name)
        if identity is not None:
            grouped[path.parent].append((path, identity))
        elif SOURCE_LIKE_FILENAME_RE.match(path.name):
            raise ValueError(f"疑似 AIS 來源檔無法解析，不能從全量計畫中略過：{path}")
    jobs: list[dict[str, object]] = []
    for folder, entries in grouped.items():
        port_months = {(identity[0], identity[1][:7]) for _path, identity in entries}
        if len(port_months) != 1:
            raise ValueError(f"來源資料夾混合港別或月份：{folder} / {sorted(port_months)}")
        port, month = next(iter(port_months))
        year, month_number = map(int, month.split("-"))
        days = {identity[1] for _path, identity in entries}
        jobs.append(
            {
                "folder": folder,
                "port": port,
                "year": year,
                "month": month_number,
                "physical_files": len(entries),
                "tolerant_logical_days": len(days),
                "logical_days": sorted(days),
                "source_files": sorted(path.name for path, _identity in entries),
            }
        )
    jobs = sorted(jobs, key=lambda job: (str(job["port"]), int(job["year"]), int(job["month"])))
    tokens = [job_token(job) for job in jobs]
    if len(tokens) != len(set(tokens)):
        raise ValueError("來源包含相同港別／月份的多個工作資料夾，會碰撞同一輸出；請先確認來源布局。")
    return jobs


def run_job(job: dict[str, object], output_root: Path, workers: int) -> dict[str, object]:
    port = str(job["port"])
    year = int(job["year"])
    month = int(job["month"])
    folder = Path(job["folder"])
    token = f"{port}_{year}{month:02d}"
    job_root = output_root / token
    modern = job_root / default_output_filename(port, year, month)
    delivery = job_root / "正式交付" / token
    catalog = detect_source_catalog(folder)
    detected = next(
        (
            count
            for detected_year, detected_month, count in catalog.get(port, [])
            if (detected_year, detected_month) == (year, month)
        ),
        0,
    )
    events_path = job_root / "events.jsonl"
    job_root.mkdir(parents=True, exist_ok=True)

    def callback(payload: dict[str, object]) -> None:
        with events_path.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(payload, ensure_ascii=False, default=str) + "\n")

    started = time.perf_counter()
    record: dict[str, object] = {
        **serializable_job(job),
        "current_parser_logical_days": detected,
        "modern": str(modern),
        "delivery": str(delivery),
    }
    try:
        run_pipeline(
            AppConfig(
                input_dir=folder,
                output_path=modern,
                port=port,
                year=year,
                month=month,
                delivery_dir=delivery,
                workers=workers,
                overwrite=True,
            ),
            callback=callback,
        )
        record["status"] = "PASS"
    except Exception as error:
        record.update(
            {
                "status": "FAIL",
                "error": f"{type(error).__name__}: {error}",
                "traceback": traceback.format_exc(),
            }
        )
    record["elapsed_seconds"] = round(time.perf_counter() - started, 3)
    record["finished_at"] = datetime.now().astimezone().isoformat()
    return record


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("source_root", type=Path)
    parser.add_argument("output_root", type=Path)
    parser.add_argument("--workers", type=int, default=3)
    parser.add_argument("--only", action="append", default=[])
    arguments = parser.parse_args(argv)
    source_resolved, output_resolved = arguments.source_root.resolve(), arguments.output_root.resolve()
    if (source_resolved == output_resolved or source_resolved in output_resolved.parents
            or output_resolved in source_resolved.parents):
        raise ValueError("回歸輸出與原始來源目錄必須完全分離，不能互為父子目錄。")
    report_path = arguments.output_root / "run_report.json"
    protected_sources = list(arguments.source_root.rglob("*.xlsx"))
    guard_output(report_path, [arguments.source_root], protected_sources)
    arguments.output_root.mkdir(parents=True, exist_ok=True)
    records: list[dict[str, object]] = []
    report: dict[str, object] = {"source_root": str(arguments.source_root.resolve()),
        "started_at": datetime.now().astimezone().isoformat(), "status": "IN_PROGRESS",
        "scope": "SUBSET" if arguments.only else "FULL", "planned_jobs": [],
        "expected_job_count": 0, "jobs": records}
    # Reset the status before discovery too; an old COMPLETE report cannot stand
    # in for a new attempt that fails during source discovery or its first job.
    atomic_json(report_path, report)
    try:
        jobs = discover_jobs(arguments.source_root)
        report["available_job_count"] = len(jobs)
        if arguments.only:
            allowed = set(arguments.only)
            missing = allowed - {job_token(job) for job in jobs}
            if missing:
                raise ValueError(f"--only 未找到工作：{sorted(missing)}")
            jobs = [job for job in jobs if job_token(job) in allowed]
        if not jobs:
            raise ValueError("全量回歸沒有適用工作，不能回報成功。")
        planned = [serializable_job(job) for job in jobs]
        report.update(planned_jobs=planned, expected_job_count=len(planned),
            plan_sha256=hashlib.sha256(json.dumps(planned, ensure_ascii=False,
                sort_keys=True).encode("utf-8")).hexdigest())
        atomic_json(report_path, report)
        for position, job in enumerate(jobs, start=1):
            token = job_token(job)
            report["active_job"] = token
            atomic_json(report_path, report)
            print(f"[{position}/{len(jobs)}] {token}", flush=True)
            records.append(run_job(job, arguments.output_root, arguments.workers))
            atomic_json(report_path, report)
            print(json.dumps(records[-1], ensure_ascii=False, indent=2), flush=True)
    except BaseException as error:
        report.update(status="INTERRUPTED" if isinstance(error, (KeyboardInterrupt, SystemExit)) else "FAILED",
                      error=f"{type(error).__name__}: {error}",
                      finished_at=datetime.now().astimezone().isoformat())
        atomic_json(report_path, report)
        raise
    report.pop("active_job", None)
    report.update(status="COMPLETE", finished_at=datetime.now().astimezone().isoformat())
    atomic_json(report_path, report)
    return 1 if any(record["status"] != "PASS" for record in records) else 0


if __name__ == "__main__":
    raise SystemExit(main())
