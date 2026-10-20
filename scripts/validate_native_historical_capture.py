"""Recheck every staged Historical native source/helper/series/geometry observation.

This accepts only current hash-bound independent period contracts; no blanket NA,
missing-as-zero tolerance, crop or original/reference write is permitted.
"""
from __future__ import annotations

import hashlib
import json
import math
from pathlib import Path
import re


def canonical_native_formula(formula):
    """Normalize only Excel-resolved MAXIFS/STDEV.S tokens outside literals."""
    parts = re.split(r'("(?:[^"]|"")*")', formula)
    for index in range(0, len(parts), 2):
        parts[index] = re.sub(r"(?<![A-Za-z0-9_.])_xlfn\.(MAXIFS|STDEV\.S)\(", r"\1(", parts[index])
    return "".join(parts)


def digest(path):
    with Path(path).open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest().upper()


def value(actual, expected, *, primary=False):
    if expected is None:
        if actual is not None and not (primary and actual == ""):
            raise ValueError("Historical missing data became numeric/text")
        return
    if any(isinstance(v, bool) or not isinstance(v, (int, float)) or not math.isfinite(v) or v < 0
           for v in (actual, expected)):
        raise ValueError("Historical native number is invalid")
    if ((expected == 0) != (actual == 0)
            or abs(actual - expected) > 1e-12 * max(1, abs(expected))):
        raise ValueError("Historical native number differs from independent matrix")


def point_geometry(point, plot, expected, chart, series, index, count, axis_min, axis_max):
    fields = ("left", "top", "width", "height")
    if any(isinstance(point.get(field), bool) or not isinstance(point.get(field), (int, float))
           or not math.isfinite(point[field]) for field in fields):
        raise ValueError("Historical geometry unavailable/nonfinite")
    if (any(isinstance(v, bool) or not isinstance(v, (int, float)) or not math.isfinite(v)
            for v in (*plot.values(), axis_min, axis_max))
            or min(point["width"], point["height"]) < 0 or min(plot["width"], plot["height"]) <= 0 or axis_max <= axis_min):
        raise ValueError("Historical geometry/axis dimensions invalid")
    intersects = (point["left"]+point["width"] >= plot["left"]-.02
                  and point["left"] <= plot["left"]+plot["width"]+.02
                  and point["top"]+point["height"] >= plot["top"]-.02
                  and point["top"] <= plot["top"]+plot["height"]+.02)
    if expected is None:
        anchor_inside = (plot["left"]-.02 <= point["left"] <= plot["left"]+plot["width"]+.02
                         and plot["top"]-.02 <= point["top"] <= plot["top"]+plot["height"]+.02)
        if anchor_inside:
            raise ValueError("Historical missing point anchor intersects actual plot")
        return
    if not intersects or not axis_min <= expected <= axis_max:
        raise ValueError("Historical numeric point is outside plot/axis")
    ratio = (expected-axis_min)/(axis_max-axis_min)
    if chart == 1:
        angle = 2*math.pi*(index-1)/count
        radius = min(plot["width"], plot["height"])/2*ratio
        x = plot["left"]+plot["width"]/2+radius*math.sin(angle)
        y = plot["top"]+plot["height"]/2-radius*math.cos(angle)
        if (abs(point["left"]+point["width"]/2-x) > .02
                or abs(point["top"]+point["height"]/2-y) > .02):
            raise ValueError("Historical radar point does not represent independent direction/value")
    else:
        y = plot["top"]+plot["height"]*(1-ratio)
        observed_y = point["top"]+(point["height"]/2 if series == 3 else 0)
        if abs(observed_y-y) > .02 or (series < 3 and abs(point["height"]-(plot["top"]+plot["height"]-y)) > .02):
            raise ValueError("Historical statistic point geometry differs from independent value")


def validate_stage(gate, contract):
    if (gate.get("status"), gate.get("independent_contract_status"), gate.get("missing_gate_kind")) != ("PASS", "PASS", "Historical/radar") or gate.get("error"):
        raise ValueError("Historical native gate incomplete")
    if (gate["role"], gate["scope"]) != (contract["role"], contract["scope"]):
        raise ValueError("Historical native role/scope differs")
    sources = contract["source_cells"]
    observed = gate["source_observations"]
    if len(sources) != len(observed) or gate["source_cells"] != len(sources):
        raise ValueError("Historical source-cell native coverage incomplete")
    for expected, actual in zip(sources, observed):
        if (not isinstance(actual["has_formula"], bool)
                or (actual["sheet"], actual["cell"], actual["has_formula"]) != (expected["sheet"], expected["cell"], expected["formula"] is not None)):
            raise ValueError("Historical native source/formula mapping differs")
        if (expected["formula"] is not None
                and canonical_native_formula(actual["formula"]) != canonical_native_formula(expected["formula"])):
            raise ValueError("Historical native source formula differs")
        value(actual["value"], expected["expected_value"], primary=True)
    helpers = contract["radar_helpers"]
    observed = gate["helper_observations"]
    if len(helpers) != len(observed) or gate["radar_helper_cells"] != len(helpers):
        raise ValueError("Historical native radar-helper coverage incomplete")
    allowed = {helper["cell"] for helper in helpers if helper["expected_missing"]}
    if len(gate["intentional_error_addresses"]) != len(allowed) or set(gate["intentional_error_addresses"]) != allowed:
        raise ValueError("Historical native exact intentionalNA scope differs")
    for expected, actual in zip(helpers, observed):
        if (actual["cell"], actual["source"], actual["formula"], actual["expected_missing"]) != (expected["cell"], expected["source"], expected["formula"], expected["expected_missing"]):
            raise ValueError("Historical native helper formula/source/missing mapping differs")
        value(actual["source_value"], expected["expected_value"], primary=True)
        if expected["expected_missing"]:
            if actual["source_count"] != 0 or actual["helper_value"] != -2146826246:
                raise ValueError("Historical deliberateNA lacks exact source/ISNA proof")
        else:
            if actual["source_count"] != 1:
                raise ValueError("Historical genuine numeric source has wrong COUNT")
            value(actual["helper_value"], expected["expected_value"])
    plan = [(1, index, series) for index, series in enumerate(contract["radar_series"], 1)]
    plan += [(2, index, series) for index, series in enumerate(contract["statistic_series"], 1)]
    if len(gate["series_observations"]) != len(plan):
        raise ValueError("Historical native series coverage incomplete")
    count = missing = zeros = 0
    for actual, (chart, index, expected) in zip(gate["series_observations"], plan):
        labels = (contract["direction_labels"] if chart == 1 else
                  [contract["direction_labels"][direction] for direction in contract["statistic_order"]])
        values = expected["expected_values"]
        absolute = lambda text: re.sub(r"([A-Z]+)([0-9]+)", r"$\1$\2", text)
        formula = f'=SERIES("{expected["name"]}",{absolute(expected["categories"])},{absolute(expected["values"])},{index})'
        if (actual["chart_index"], actual["series_index"], actual["name"], actual["formula"], actual["categories"], actual["point_count"]) != (chart, index, expected["name"], formula, labels, len(values)):
            raise ValueError("Historical native series source/category/point mapping differs")
        if len(actual["values"]) != len(values) or len(actual["points"]) != len(values):
            raise ValueError("Historical actual per-point coverage incomplete")
        for point_index, (seen, wanted, geometry) in enumerate(zip(actual["values"], values, actual["points"]), 1):
            value(seen, wanted)
            if geometry["index"] != point_index:
                raise ValueError("Historical native point geometry order differs")
            point_geometry(geometry, actual["plot"], wanted, chart, index, point_index, len(values), actual["axis_minimum"], actual["axis_maximum"])
        expected_missing = sum(item is None for item in values)
        expected_zero = sum(item == 0 for item in values if item is not None)
        if (actual["missing_native_null_points"], actual["genuine_zero_points"]) != (expected_missing, expected_zero):
            raise ValueError("Historical per-series missing/genuine0 counts differ")
        count += len(values); missing += expected_missing; zeros += expected_zero
    if (gate["series_points"], gate["missing_points"], gate["genuine_zero_points"]) != (count, missing, zeros):
        raise ValueError("Historical complete native point totals differ")
    return allowed


def validated_historical_gate(book):
    gate = book.get("historical_chart_missing")
    if gate is None:
        return set()
    report_path = Path(gate["contract_report"])
    if digest(report_path) != gate["contract_report_sha256"].upper():
        raise ValueError("Historical independent contract report changed")
    report = json.loads(report_path.read_text(encoding="utf-8-sig"))
    contracts = report.get("historical_chart_contracts", [])
    if not contracts:
        contracts = [contract for job in report["jobs"] if job.get("pass") is True for contract in job["historical_chart_contracts"]]
    path = Path(book["path"]).resolve(strict=True)
    matched = [contract for contract in contracts if Path(contract["workbook"]).resolve(strict=True) == path]
    if len(matched) != 1:
        raise ValueError("Historical independent workbook contract missing/duplicated")
    contract = matched[0]
    if (contract["schema"], contract["state"]) != (1, "PASS_SERIALIZATION_PREREQUISITES") or contract["workbook_sha256"].upper() != digest(path) or gate["workbook_sha256"].upper() != digest(path):
        raise ValueError("Historical contract schema/state/current workbook hash invalid")
    if not contract["cache_sources"]:
        raise ValueError("Historical current independent cache bindings absent")
    for source in contract["cache_sources"]:
        if digest(source["path"]) != source["sha256"].upper():
            raise ValueError("Historical independent cache source changed")
    allowed = validate_stage(gate, contract)
    stages = book["historical_chart_stages"]
    if [stage["stage"] for stage in stages] != ["AFTER_FULL_REBUILD_BEFORE_EXPORT", "AFTER_NATIVE_EXPORT", "AFTER_EXPLICIT_REFRESH"]:
        raise ValueError("Historical actual staged native coverage incomplete")
    for stage in stages:
        if validate_stage(stage["gate"], contract) != allowed:
            raise ValueError("Historical native missingness changes across export/refresh")
    return allowed
