from __future__ import annotations

import argparse
import json
import math
from dataclasses import replace
from datetime import date
from pathlib import Path

import openpyxl

import ais_monthly_app as app


def read_grid(path: Path) -> list[list[float | None]]:
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook["總表"]
        return [
            [
                None if value in (None, "") else float(value)
                for value in row
            ]
            for row in sheet.iter_rows(
                min_row=2, max_row=31, min_col=2, max_col=33, values_only=True
            )
        ]
    finally:
        workbook.close()


def compare_grid(left: list[list[float | None]], right: list[list[float | None]]) -> dict[str, object]:
    for name, grid in (("reference", left), ("generated", right)):
        if len(grid) != 30 or any(len(row) != 32 for row in grid):
            raise ValueError(f"{name} grid must contain all 30 days and 32 directions")
    mismatches: list[dict[str, object]] = []
    for day_offset, (left_row, right_row) in enumerate(zip(left, right), start=1):
        for direction_index, (expected, actual) in enumerate(zip(left_row, right_row)):
            same = expected is actual or (
                expected is not None
                and actual is not None
                and math.isclose(expected, actual, rel_tol=0.0, abs_tol=1e-9)
            )
            if not same:
                mismatches.append(
                    {
                        "day": day_offset,
                        "direction_index": direction_index,
                        "expected": expected,
                        "actual": actual,
                    }
                )
    return {
        "cells": 30 * 32,
        "mismatch_count": len(mismatches),
        "mismatch_sample": mismatches[:20],
        "pass": not mismatches,
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Reproduce the reviewed June delivery from raw-candidate evidence."
    )
    parser.add_argument("cache_directory", type=Path)
    parser.add_argument("evidence_json", type=Path)
    parser.add_argument("reference_directory", type=Path)
    parser.add_argument("source_directory", type=Path)
    parser.add_argument("output_directory", type=Path)
    parser.add_argument("--summaries-only", action="store_true")
    parser.add_argument("--report", type=Path, required=True)
    arguments = parser.parse_args()

    day_results = []
    for day_number in range(1, 31):
        payload = json.loads(
            (arguments.cache_directory / f"202606{day_number:02d}.json").read_text(
                encoding="utf-8"
            )
        )
        day_results.append(app._day_result_from_dict(payload["result"]))

    cache_suffix = "_KLNG_2026_06_cache"
    cache_name = arguments.cache_directory.name
    if not cache_name.startswith(".") or not cache_name.endswith(cache_suffix):
        raise ValueError("Unexpected KLNG June cache directory name.")
    analysis_path = arguments.cache_directory.parent / (
        cache_name[1 : -len(cache_suffix)] + ".xlsx"
    )
    config = app.AppConfig(
        input_dir=arguments.source_directory,
        # The delivery full workbooks read normalized rows from the analysis
        # spool directory derived from this path.  Keep the already-audited
        # baseline analysis location and publish only to output_directory.
        output_path=analysis_path,
        delivery_dir=arguments.output_directory,
        port="KLNG",
        year=2026,
        month=6,
        overwrite=True,
        workers=1,
    )
    config.validate()
    default_snapshot = app.build_default_decision_snapshot(config, day_results)
    evidence_payload = json.loads(arguments.evidence_json.read_text(encoding="utf-8"))
    if not evidence_payload.get("pass") or evidence_payload.get("matched") != 1260:
        raise ValueError("Golden candidate evidence is incomplete or did not pass.")

    candidate_ids = {
        (
            item["period"],
            date.fromisoformat(item["date"]),
            app.DIRECTION_ORDER[int(item["direction_index"])],
        ): item["candidate_id"]
        for item in evidence_payload["matched_evidence"]
    }
    reviewed = []
    for decision in default_snapshot.decisions:
        if decision.scope in (app.PERIOD_A, app.PERIOD_B):
            candidate_id = candidate_ids.get(
                (decision.scope, decision.day, decision.direction)
            )
            if candidate_id is None:
                raise ValueError(
                    f"Missing reviewed candidate: {decision.scope} "
                    f"{decision.day} {decision.direction}"
                )
            reviewed.append(replace(decision, candidate_id=candidate_id))
        else:
            reviewed.append(decision)
    snapshot = app.DecisionSnapshot(
        port=default_snapshot.port,
        year=default_snapshot.year,
        month=default_snapshot.month,
        source_manifest_hash=default_snapshot.source_manifest_hash,
        decisions=tuple(reviewed),
    )
    app.validate_decision_snapshot_contract(snapshot, config, day_results)
    if arguments.summaries_only:
        arguments.output_directory.mkdir(parents=True, exist_ok=True)
        period_a = app._period_values(config, snapshot, day_results, app.PERIOD_A)
        period_b = app._period_values(config, snapshot, day_results, app.PERIOD_B)
        integrated = app._integrated_period_values(day_results, period_a, period_b)
        summary_outputs = {
            "period_a_summary": arguments.output_directory / "KLNG 6月 32方位數值_2 總表.xlsx",
            "period_b_summary": arguments.output_directory / "KLNG 6月 32方位數值 總表.xlsx",
            "integrated": arguments.output_directory / "KLNG_6月_32方位_每日較大值整合總表.xlsx",
        }
        app._write_summary_workbook(
            summary_outputs["period_a_summary"],
            config,
            day_results,
            period_a,
            artifact_role="summary_a",
            scope_title="Historical Period A 00:00–11:59",
        )
        app._write_summary_workbook(
            summary_outputs["period_b_summary"],
            config,
            day_results,
            period_b,
            artifact_role="summary_b",
            scope_title="Historical Period B 12:00–23:59",
        )
        app._write_summary_workbook(
            summary_outputs["integrated"],
            config,
            day_results,
            integrated,
            artifact_role="integrated",
            scope_title="Historical Period A/B 整合 MAX",
            integrated_inputs=(period_a, period_b),
        )
        paths = None
    else:
        paths = app.write_delivery_workbooks(
            config, day_results, snapshot, root=arguments.output_directory
        )

    references = {
        "period_a_full": arguments.reference_directory / "KLNG 6月 32方位數值_2.xlsx",
        "period_b_full": arguments.reference_directory / "KLNG 6月 32方位數值.xlsx",
        "period_a_summary": arguments.reference_directory / "KLNG 6月 32方位數值_2 總表.xlsx",
        "period_b_summary": arguments.reference_directory / "KLNG 6月 32方位數值 總表.xlsx",
        "integrated": arguments.reference_directory / "KLNG_6月_32方位_每日較大值整合總表.xlsx",
    }
    outputs = summary_outputs if arguments.summaries_only else {
        "period_a_full": paths.period_a_full,
        "period_b_full": paths.period_b_full,
        "period_a_summary": paths.period_a_summary,
        "period_b_summary": paths.period_b_summary,
        "integrated": paths.integrated_summary,
    }
    comparisons = {
        name: compare_grid(read_grid(references[name]), read_grid(path))
        for name, path in outputs.items()
    }
    report = {
        "candidate_evidence": {
            "matched": evidence_payload["matched"],
            "missing_count": evidence_payload["missing_count"],
            "displayed_candidate_misses": evidence_payload[
                "displayed_candidate_misses"
            ],
        },
        "outputs": {name: str(path) for name, path in outputs.items()},
        "comparisons": comparisons,
        "pass": all(result["pass"] for result in comparisons.values()),
    }
    arguments.report.parent.mkdir(parents=True, exist_ok=True)
    arguments.report.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(
        json.dumps(
            {
                "pass": report["pass"],
                "mismatch_counts": {
                    name: result["mismatch_count"]
                    for name, result in comparisons.items()
                },
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
