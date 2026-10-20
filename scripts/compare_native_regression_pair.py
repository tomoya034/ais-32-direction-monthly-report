"""Exhaustive current-file/geometry/pixel comparison of one six-book month.

Read-only native captures are necessary, not sufficient, AIS correctness proof.
This comparator does not replace independent source/numeric/mapping audits.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
from pathlib import Path
try:
    from scripts.validate_native_historical_capture import validated_historical_gate
except ModuleNotFoundError as error:
    if error.name != "scripts":
        raise
    from validate_native_historical_capture import validated_historical_gate

GEOMETRY_FIELDS = (
    "chart_type", "chart_style", "title", "left_points", "top_points",
    "width_points", "height_points", "plot_left_points", "plot_top_points",
    "plot_width_points", "plot_height_points", "series_count",
)


def sha256(path: Path) -> str:
    with path.open("rb") as stream:
        digest = hashlib.file_digest(stream, "sha256")
    return digest.hexdigest().upper()


def _native_value(actual, expected):
    if expected is None:
        if actual is not None:
            raise ValueError("Native missing series/source became a number or text")
        return
    for value in (actual, expected):
        if (isinstance(value, bool) or not isinstance(value, (int, float))
                or not math.isfinite(value) or value < 0):
            raise ValueError("Native numeric series/source is an error or invalid value")
    if ((expected == 0) != (actual == 0)
            or abs(actual - expected) > 1e-12 * max(1, abs(expected))):
        raise ValueError("Native numeric series/source differs from independent expected value")


def validated_missing_gate(book: dict) -> set[str]:
    gate = book.get("modern_chart_missing")
    if gate is None:
        return set()
    if (gate.get("status"), gate.get("independent_contract_status")) != ("PASS", "PASS") or gate.get("error"):
        raise ValueError("Incomplete or failed native Modern missing gate")
    report_path = Path(gate["contract_report"])
    if sha256(report_path) != gate["contract_report_sha256"].upper():
        raise ValueError("Independent Modern contract report changed")
    report = json.loads(report_path.read_text(encoding="utf-8-sig"))
    if "modern_chart_missing_contract" in report:
        contracts = [report["modern_chart_missing_contract"]]
    else:
        contracts = [job["modern_chart_missing_contract"] for job in report["jobs"]
                     if job.get("pass") is True and "modern_chart_missing_contract" in job]
    path = Path(book["path"]).resolve(strict=True)
    matched = [item for item in contracts if Path(item["workbook"]).resolve(strict=True) == path]
    if len(matched) != 1:
        raise ValueError("Missing or duplicate independent Modern contract")
    contract = matched[0]
    if (contract.get("schema"), contract.get("role"), contract.get("state")) != (1, "Modern/default", "PASS_SERIALIZATION_PREREQUISITES"):
        raise ValueError("Independent Modern contract schema/state failed")
    if contract["workbook_sha256"].upper() != sha256(path) or gate["workbook_sha256"].upper() != sha256(path):
        raise ValueError("Native missing gate is not bound to the current workbook")
    if not contract["cache_sources"]:
        raise ValueError("Native missing gate lacks independent source bindings")
    for source in contract["cache_sources"]:
        if sha256(Path(source["path"])) != source["sha256"].upper():
            raise ValueError("Independent native source cache changed")
    helpers = contract["helpers"] + contract["radar_helpers"]
    allowed = {item["cell"] for item in helpers if item["expected_missing"] is True}
    addresses = gate["intentional_error_addresses"]
    if len(set(addresses)) != len(addresses) or set(addresses) != allowed:
        raise ValueError("Native intentional marker allowlist differs from independent helpers")
    observations = gate["helper_observations"]
    if len(observations) != len(helpers) or len({item["cell"] for item in observations}) != len(helpers):
        raise ValueError("Native helper census is incomplete or duplicated")
    expected_by_cell = {item["cell"]: item for item in helpers}
    if set(expected_by_cell) != {item["cell"] for item in observations}:
        raise ValueError("Native helpers have shifted/extra scope")
    for actual in observations:
        expected = expected_by_cell[actual["cell"]]
        if (actual["source"], actual["formula"], actual["expected_missing"]) != (expected["source"], expected["formula"], expected["expected_missing"]):
            raise ValueError("Native helper reference/formula/missingness differs")
        if expected["expected_missing"]:
            if (actual["source_count"] != 0 or actual["source_value"] not in (None, "")
                    or actual["helper_value"] != -2146826246):
                raise ValueError("Native deliberateNA lacks missing source/exactNA proof")
        else:
            if actual["source_count"] != 1:
                raise ValueError("Native numeric source count is not1")
            _native_value(actual["source_value"], expected["expected_value"])
            _native_value(actual["helper_value"], expected["expected_value"])
    matrix = contract["expected_matrix"]
    indices = contract["included_indices"]
    labels = contract["direction_labels"]
    days = contract["logical_days"]
    if (len(matrix) != len(days) or len(labels) != 32
            or any(len(row) != 32 for row in matrix)
            or len(indices) != len(set(indices))):
        raise ValueError("Invalid independent native matrix dimensions")
    if (gate["primary_daily_cells"], gate["statistic_helper_cells"], gate["radar_helper_cells"]) != (32*len(days), 3*len(indices), 32*len(days)):
        raise ValueError("Native primary/helper matrix coverage is incomplete")
    radar_days = [index for index, row in enumerate(matrix) if any(value is not None for value in row)]
    empty = not radar_days
    if empty:
        radar_days = [0]
    series = gate["series_observations"]
    if len(series) != len(radar_days) + 3:
        raise ValueError("Native series census differs from independent matrix")
    expected_series = []
    for position, day_index in enumerate(radar_days, 1):
        row = day_index + 5
        month, day = map(int, days[day_index].split("-")[1:])
        name = "本時段無有效資料" if empty else f"{month}/{day}"
        formula = f'=SERIES("{name}",總表!$B$4:$AG$4,總表!$AO${row}:$BT${row},{position})'
        expected_series.append((1, position, name, formula, matrix[day_index], labels))
    for position, (name, column) in enumerate(zip(("最遠", "最近", "平均"), ("AJ", "AK", "AL")), 1):
        last = len(indices) + 4
        values = [contract["helpers"][3*offset+position-1]["expected_value"] for offset in range(len(indices))]
        formula = f'=SERIES("{name}",總表!$AI$5:$AI${last},總表!${column}$5:${column}${last},{position})'
        expected_series.append((2, position, name, formula, values, [labels[index] for index in indices]))
    points = zeros = 0
    for actual, expected in zip(series, expected_series):
        chart, position, name, formula, values, categories = expected
        if (actual["chart_index"], actual["series_index"], actual["name"], actual["formula"], actual["point_count"], actual["categories"]) != (chart, position, name, formula, len(values), categories):
            raise ValueError("Native series source/category/order/count drift")
        if len(actual["values"]) != len(values):
            raise ValueError("Native series value count drift")
        for actual_value, expected_value in zip(actual["values"], values):
            _native_value(actual_value, expected_value)
        missing = sum(value is None for value in values)
        genuine_zero = sum(value == 0 for value in values if value is not None)
        if (actual["missing_native_null_points"], actual["genuine_zero_points"]) != (missing, genuine_zero):
            raise ValueError("Native missing/genuine-zero census drift")
        points += len(values)
        zeros += genuine_zero
    if (gate["series_points"], gate["genuine_zero_points"]) != (points, zeros):
        raise ValueError("Native total point/zero census drift")
    return allowed


def validated_capture(report: dict) -> dict[str, dict]:
    if report.get("state") != "COMPLETE":
        raise ValueError("Incomplete native capture cannot be acceptance evidence")
    if (report.get("expected_workbook_count"), report.get("workbook_count"),
        report.get("pass_count"), report.get("fail_count")) != (6, 6, 6, 0):
        raise ValueError("Six-book native acceptance contract failed")
    root = Path(report["root"]).resolve(strict=True)
    files = {p.resolve(strict=True).relative_to(root).as_posix(): p
             for p in root.rglob("*.xlsx")}
    if len(files) != 6 or any(p.name.startswith("~$") for p in files.values()):
        raise ValueError("Current output census is not exactly six final workbooks")
    entries = {}
    for book in report["workbooks"]:
        path = Path(book["path"]).resolve(strict=True)
        relative = path.relative_to(root).as_posix()
        if relative in entries:
            raise ValueError("Duplicate native workbook capture")
        if (book.get("status") != "PASS" or not book.get("opened_read_only")
                or not book.get("file_unchanged") or book.get("error")
                or book.get("formula_errors")):
            raise ValueError("Failed/unknown native workbook capture")
        if not (book.get("sha256_before") == book.get("sha256_after") == sha256(path)):
            raise ValueError("Current workbook changed after native capture")
        checks = book["formula_checks"]
        allowed_missing = validated_missing_gate(book) | validated_historical_gate(book)
        if ([item["sheet"] for item in checks] != book["sheets"]
                or len(set(book["sheets"])) != len(book["sheets"])):
            raise ValueError("Formula coverage does not cover every sheet exactly once")
        for check in checks:
            result = check["result"]
            if result["status"] not in {"VERIFIED_NO_FORMULAS", "VERIFIED_ZERO_ERRORS", "VERIFIED_INTENTIONAL_CHART_MISSING"}:
                raise ValueError("Unknown/failed native formula scan")
            expected_markers = allowed_missing if check["sheet"] == "總表" else set()
            actual_markers = result.get("intentional_chart_missing_addresses", [])
            if (len(set(actual_markers)) != len(actual_markers) or set(actual_markers) != expected_markers
                    or result.get("intentional_chart_missing_count", 0) != len(expected_markers)
                    or result.get("unexpected_error_count", result.get("error_count")) != 0
                    or result.get("error_count") != len(expected_markers) or result.get("error")
                    or (bool(expected_markers) != (result["status"] == "VERIFIED_INTENTIONAL_CHART_MISSING"))):
                raise ValueError("Formula errors present in native capture")
            formula_count = result.get("formula_cell_count")
            if (isinstance(formula_count, bool) or not isinstance(formula_count, int)
                    or formula_count < 0
                    or (result["status"] in {"VERIFIED_ZERO_ERRORS", "VERIFIED_INTENTIONAL_CHART_MISSING"} and formula_count == 0)
                    or (result["status"] == "VERIFIED_NO_FORMULAS" and formula_count != 0)):
                raise ValueError("Inconsistent formula coverage/count verdict")
        chart_map = {}
        for chart in book["charts"]:
            key = (chart["sheet"], chart["index"])
            if key in chart_map:
                raise ValueError("Duplicate native chart capture")
            validation = chart["png_validation"]
            png = Path(chart["png"])
            if not chart["png_exported"] or not validation.get("decoded"):
                raise ValueError("Native chart PNG not exported/decoded")
            if validation["status"] != "VALID_PNG" or validation["sha256"] != sha256(png):
                raise ValueError("Chart PNG changed after native capture")
            if any(field not in chart for field in GEOMETRY_FIELDS):
                raise ValueError("Incomplete native geometry inventory")
            for field in GEOMETRY_FIELDS:
                if field.endswith("_points") and not math.isfinite(chart[field]):
                    raise ValueError("Nonfinite native geometry cannot pass")
            chart_map[key] = chart
        if len(chart_map) != 2:
            raise ValueError("Each workbook must have exactly two chart captures")
        entries[relative] = {"book": book, "charts": chart_map}
    if set(entries) != set(files):
        raise ValueError("Native workbook capture omitted or added a current output")
    if report.get("historical_contract_report"):
        historical = [item["book"]["historical_chart_missing"] for item in entries.values()
                      if item["book"].get("historical_chart_missing") is not None]
        if len(historical) != 5 or len({item["role"] for item in historical}) != 5:
            raise ValueError("Historical native mandatory five-role coverage incomplete")
    return entries


def compare_png(first: Path, repeat: Path) -> dict:
    from PIL import Image, ImageChops
    with Image.open(first) as left, Image.open(repeat) as right:
        left.load(); right.load()
        if left.size != right.size:
            return {"status": "DIFFERENT_DIMENSIONS", "first_size": left.size,
                    "repeat_size": right.size, "changed_pixels": None}
        # RGBA getbbox alone ignores RGB changes when the alpha delta is zero.
        # Combine ALL channel deltas, without resizing/cropping/compositing.
        difference = ImageChops.difference(left.convert("RGBA"), right.convert("RGBA"))
        channels = difference.split()
        maximum = channels[0]
        for channel in channels[1:]:
            maximum = ImageChops.lighter(maximum, channel)
        pixels = left.width * left.height
        changed = pixels - maximum.histogram()[0]
        return {"status": "EXACT_EQUAL" if changed == 0 else "PIXEL_DIFFERENCE",
                "width_pixels": left.width, "height_pixels": left.height,
                "changed_pixels": changed, "difference_bounds": maximum.getbbox()}


def compare(first: dict, repeat: dict) -> dict:
    left, right = validated_capture(first), validated_capture(repeat)
    if set(left) != set(right):
        raise ValueError("Repeated run has a different workbook role/path census")
    pairs = []
    for relative in sorted(left):
        a, b = left[relative]["charts"], right[relative]["charts"]
        if set(a) != set(b):
            raise ValueError("Repeated run chart worksheet/index census differs")
        for key in sorted(a):
            one, two = a[key], b[key]
            differences = {field: {"first": one[field], "repeat": two[field]}
                           for field in GEOMETRY_FIELDS if one[field] != two[field]}
            pixels = compare_png(Path(one["png"]), Path(two["png"]))
            pairs.append({"workbook_relative_path": relative, "sheet": key[0],
                          "chart_index": key[1], "native_geometry_differences": differences,
                          "pixels": pixels,
                          "status": "PASS" if not differences and pixels["status"] == "EXACT_EQUAL" else "FAIL"})
    return {"schema": "AIS_NATIVE_MONTH_REPEAT_1", "first_root": first["root"],
            "repeat_root": repeat["root"], "workbooks_each": 6,
            "chart_pairs": len(pairs), "pass_count": sum(p["status"] == "PASS" for p in pairs),
            "status": "PASS" if all(p["status"] == "PASS" for p in pairs) else "FAIL",
            "comparison_contract": "Every current workbook hash, every sheet formula verdict, exact native geometry and decoded RGBA pixels; no crop/resize",
            "independent_numeric_source_manual_parity": "UNVERIFIED by this visual comparator; separate root-owned evidence required",
            "pairs": pairs}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--first-report", required=True, type=Path)
    parser.add_argument("--repeat-report", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    if args.output.exists():
        raise ValueError("Evidence output already exists")
    first = json.loads(args.first_report.read_text(encoding="utf-8-sig"))
    repeat = json.loads(args.repeat_report.read_text(encoding="utf-8-sig"))
    result = compare(first, repeat)
    result["source_reports"] = [{"path": str(p), "sha256": sha256(p)}
                                for p in (args.first_report, args.repeat_report)]
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({k: result[k] for k in ("status", "workbooks_each", "chart_pairs", "pass_count")}, ensure_ascii=False))
    return 0 if result["status"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
