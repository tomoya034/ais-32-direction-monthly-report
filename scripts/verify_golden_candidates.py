from __future__ import annotations

import argparse
import json
import math
from pathlib import Path

import openpyxl

from ais_monthly_app import (
    DIRECTION_ORDER,
    HISTORICAL_MESSAGE_TYPES,
    PERIOD_A,
    PERIOD_B,
    _candidate_id,
    _day_result_from_dict,
    degree_to_direction_index,
    read_normalized_spool,
)


def golden_candidates(path: Path, month: int) -> dict[tuple[int, int], tuple[float, float]]:
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        output: dict[tuple[int, int], tuple[float, float]] = {}
        for day_number in range(1, 31):
            sheet = workbook[f"{month}月{day_number}日"]
            for distance, bearing, direction in sheet.iter_rows(
                min_row=2, min_col=1, max_col=3, values_only=True
            ):
                if distance is None or bearing is None:
                    continue
                parsed = degree_to_direction_index(bearing)
                if parsed is None:
                    continue
                direction_index = parsed[0]
                output.setdefault(
                    (day_number, direction_index), (float(distance), float(bearing))
                )
        return output
    finally:
        workbook.close()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("spool_directory", type=Path)
    parser.add_argument("cache_directory", type=Path)
    parser.add_argument("period_a_reference", type=Path)
    parser.add_argument("period_b_reference", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    arguments = parser.parse_args()
    expected = {
        PERIOD_A: golden_candidates(arguments.period_a_reference, 6),
        PERIOD_B: golden_candidates(arguments.period_b_reference, 6),
    }
    included = tuple(range(11)) + tuple(range(22, 32))
    evidence = []
    missing = []
    displayed_misses = []
    for day_number in range(1, 31):
        cache_payload = json.loads(
            (arguments.cache_directory / f"202606{day_number:02d}.json").read_text(
                encoding="utf-8"
            )
        )
        day_result = _day_result_from_dict(cache_payload["result"])
        _counts, iterator = read_normalized_spool(
            arguments.spool_directory / f"202606{day_number:02d}.bin"
        )
        targets = {
            (period, direction_index): expected[period].get((day_number, direction_index))
            for period in (PERIOD_A, PERIOD_B)
            for direction_index in included
        }
        first_matches = {}
        multiplicities = {}
        for direction_index, record in iterator:
            key = (record.period, direction_index)
            target = targets.get(key)
            if target is None or record.msg_type not in HISTORICAL_MESSAGE_TYPES:
                continue
            distance, bearing = target
            if not (
                math.isclose(record.distance, distance, rel_tol=0.0, abs_tol=1e-9)
                and math.isclose(record.bearing, bearing, rel_tol=0.0, abs_tol=1e-9)
            ):
                continue
            multiplicities[key] = multiplicities.get(key, 0) + 1
            if key not in first_matches or record.source_key < first_matches[key].source_key:
                first_matches[key] = record
        for period in (PERIOD_A, PERIOD_B):
            for direction_index in included:
                target = expected[period].get((day_number, direction_index))
                if target is None:
                    missing.append(
                        {
                            "day": day_number,
                            "period": period,
                            "direction_index": direction_index,
                            "reason": "reference detail selection missing",
                        }
                    )
                    continue
                distance, bearing = target
                record = first_matches.get((period, direction_index))
                if record is None:
                    missing.append(
                        {
                            "day": day_number,
                            "period": period,
                            "direction_index": direction_index,
                            "distance": distance,
                            "bearing": bearing,
                            "reason": "no normalized raw candidate",
                        }
                    )
                    continue
                fragment_index = record.source_key >> 32
                fragment = day_result.fragment_info[fragment_index]
                candidate_id = _candidate_id(fragment, record)
                displayed = {
                    item.candidate_id
                    for item in day_result.period_directions[period][
                        DIRECTION_ORDER[direction_index]
                    ].candidates
                }
                if candidate_id not in displayed:
                    displayed_misses.append(
                        {
                            "day": day_number,
                            "period": period,
                            "direction_index": direction_index,
                            "candidate_id": candidate_id,
                        }
                    )
                evidence.append(
                    {
                        "date": f"202606{day_number:02d}",
                        "period": period,
                        "direction_index": direction_index,
                        "direction": DIRECTION_ORDER[direction_index],
                        "distance": record.distance,
                        "bearing": record.bearing,
                        "fragment": fragment.path.name,
                        "source_row": record.source_key & 0xFFFFFFFF,
                        "sheet": fragment.sheet,
                        "candidate_id": candidate_id,
                        "match_multiplicity": multiplicities[(period, direction_index)],
                    }
                )
    report = {
        "expected": 1260,
        "matched": len(evidence),
        "missing_count": len(missing),
        "displayed_candidate_misses": len(displayed_misses),
        "pass": len(evidence) == 1260 and not missing and not displayed_misses,
        "missing": missing,
        "displayed_misses": displayed_misses,
        "matched_evidence": evidence,
    }
    arguments.output.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(json.dumps({key: report[key] for key in (
        "expected", "matched", "missing_count", "displayed_candidate_misses", "pass"
    )}, ensure_ascii=False, indent=2))
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
