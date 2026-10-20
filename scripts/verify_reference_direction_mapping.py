from __future__ import annotations

import argparse
import json
import math
from pathlib import Path

import openpyxl

from ais_monthly_app import DIRECTION_ORDER, degree_to_direction_index


# External June workbooks predate the v1.5.1 Chinese-label correction.  These
# labels are evidence-only aliases by fixed index and are never accepted by the
# runtime decision ledger.
V150_DIRECTION_ORDER = (
    "北", "北微東", "北北東", "東北微北", "東北", "東北微東", "東北東", "東微北",
    "東", "東微南", "東南東", "東南微南", "東南", "東南微西", "南南東", "南微東",
    "南", "南微西", "南南西", "西南微南", "西南", "西南微西", "西南西", "西微南",
    "西", "西微北", "西北西", "西北微西", "西北", "西北微北", "北北西", "北微西",
)


def inspect(path: Path, limit: int) -> dict[str, object]:
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    # These particular accepted workbooks are pre-v1.5.1.  Do not merge the
    # corrected runtime labels because two old labels are textually ambiguous
    # with a different corrected index (for example 東南微南).
    label_to_index = {label: index for index, label in enumerate(V150_DIRECTION_ORDER)}
    counts = {
        "rows": 0,
        "recognized_labels": 0,
        "floor_matches": 0,
        "centered_matches": 0,
        "neither": 0,
    }
    examples: list[dict[str, object]] = []
    try:
        for sheet in workbook.worksheets:
            if sheet.title == "總表":
                continue
            for distance, bearing, label in sheet.iter_rows(
                min_row=2, min_col=1, max_col=3, values_only=True
            ):
                if distance is None or bearing is None or label is None:
                    continue
                counts["rows"] += 1
                actual = label_to_index.get(str(label))
                if actual is None:
                    if len(examples) < 20:
                        examples.append({"bearing": bearing, "label": label, "reason": "unknown label"})
                else:
                    counts["recognized_labels"] += 1
                    floor = degree_to_direction_index(bearing)
                    normalized = float(bearing) % 360.0
                    centered = int(((normalized + 5.625) % 360.0) // 11.25)
                    floor_index = floor[0] if floor is not None else None
                    if actual == floor_index:
                        counts["floor_matches"] += 1
                    if actual == centered:
                        counts["centered_matches"] += 1
                    if actual not in (floor_index, centered):
                        counts["neither"] += 1
                        if len(examples) < 20:
                            examples.append(
                                {
                                    "bearing": bearing,
                                    "label": label,
                                    "actual_index": actual,
                                    "floor_index": floor_index,
                                    "centered_index": centered,
                                }
                            )
                if counts["rows"] >= limit:
                    break
            if counts["rows"] >= limit:
                break
    finally:
        workbook.close()
    recognized = counts["recognized_labels"]
    return {
        "path": str(path),
        **counts,
        "floor_match_rate": counts["floor_matches"] / recognized if recognized else math.nan,
        "centered_match_rate": counts["centered_matches"] / recognized if recognized else math.nan,
        "examples": examples,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("workbooks", nargs="+", type=Path)
    parser.add_argument("--limit", type=int, default=100_000)
    parser.add_argument("--output", type=Path, required=True)
    arguments = parser.parse_args()
    results = [inspect(path, arguments.limit) for path in arguments.workbooks]
    payload = {
        "rule_under_test": "floor((bearing % 360) / 11.25)",
        "alternative": "floor(((bearing % 360) + 5.625) % 360 / 11.25)",
        "results": results,
        "pass": all(
            result["recognized_labels"] > 0
            and result["floor_matches"] > result["centered_matches"]
            for result in results
        ),
    }
    arguments.output.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(
        json.dumps(
            {
                "pass": payload["pass"],
                "results": [
                    {
                        key: result[key]
                        for key in (
                            "rows",
                            "recognized_labels",
                            "floor_matches",
                            "centered_matches",
                            "floor_match_rate",
                            "centered_match_rate",
                        )
                    }
                    for result in results
                ],
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0 if payload["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
