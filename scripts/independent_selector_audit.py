"""Independent numerical selector oracle over ALL normalized spool records.

No production imports. Raw-to-spool equivalence is a separate prerequisite,
proved by independent_source_audit. This oracle uses ascending sorting and
binary-search range counts, not the production rolling-window implementation.
It proves the documented automatic rule, NOT manual research decisions.
"""
from __future__ import annotations

import argparse
from bisect import bisect_left
from concurrent.futures import ProcessPoolExecutor, as_completed
from datetime import datetime, timedelta
import hashlib
import json
import math
from pathlib import Path
import struct
import time

from scripts.independent_source_audit import (
    exact_float_bin, input_stamps, plan_tasks, save_report, seal_entry,
)

LABELS = (
    "北", "北微東", "東北偏北", "東北微北", "東北", "東北微東", "東北偏東", "東微北",
    "東", "東微南", "東南偏東", "東南微東", "東南", "東南微南", "東南偏南", "南微東",
    "南", "南微西", "西南偏南", "西南微南", "西南", "西南微西", "西南偏西", "西微南",
    "西", "西微北", "西北偏西", "西北微西", "西北", "西北微北", "西北偏北", "北微西",
)
MODERN = frozenset((1, 2, 3, 18, 19))
HISTORICAL = frozenset((1, 3, 4, 18, 19))
ROW = struct.Struct("<ddQIIBq16s")
COUNTS = struct.Struct("<32Q")
FIELDS = ("selected", "selected_bearing", "selected_rank", "cluster_count",
          "over_cap_count", "over_cap_max", "selected_candidate_id")
SCHEMA = "AIS_INDEPENDENT_SELECTOR_AUDIT_3"


def oracle(records: list[tuple[float, int, float]], *, cap: float = 500.0,
           tolerance: float = 0.10, minimum: int = 3) -> dict[str, object]:
    """First three tuple fields = distance, stable source key, bearing. All rows used."""
    if not (math.isfinite(cap) and cap >= 0 and 0 < tolerance < 1 and minimum >= 2):
        raise ValueError("invalid selector settings")
    if any(not math.isfinite(row[0]) or row[0] < 0 or not math.isfinite(row[2])
           for row in records):
        raise ValueError("invalid normalized record")
    excluded = [row[0] for row in records if row[0] > cap]
    ascending = sorted((row for row in records if row[0] <= cap),
                       key=lambda row: (row[0], -row[1]))
    result: dict[str, object] = {
        "selected": None, "selected_bearing": None, "selected_rank": None,
        "selected_source_key": None,
        "cluster_count": 0, "over_cap_count": len(excluded),
        "over_cap_max": max(excluded) if excluded else None,
    }
    if not ascending:
        return result
    # The documented 1e-12 acceptance epsilon uses addition before comparison.
    # Keep that operation order in the search key for binary64 edge cases.
    shifted = [row[0] + 1e-12 for row in ascending]
    for position in range(len(ascending) - 1, minimum - 2, -1):
        ceiling, source_key, bearing = ascending[position][:3]
        lower = bisect_left(shifted, ceiling * (1.0 - tolerance), 0, position + 1)
        members = position - lower + 1
        if members >= minimum:
            result.update(selected=ceiling, selected_bearing=bearing,
                          selected_source_key=source_key,
                          selected_rank=len(ascending) - position, cluster_count=members)
            return result
    result.update(selected=ascending[-1][0], selected_bearing=ascending[-1][2],
                  selected_source_key=ascending[-1][1],
                  selected_rank=1, cluster_count=1)
    return result


def equal_field(expected, actual):
    if expected is None:
        return actual is None
    if isinstance(expected, (int, float)):
        return (isinstance(actual, (int, float)) and not isinstance(actual, bool)
                and math.isfinite(actual) and expected == actual)
    return type(expected) is type(actual) and expected == actual


def candidate_identity(fragment, record):
    """Documented provenance ID, reconstructed without production imports."""
    distance, key, bearing, second, msg, mmsi, channel = record
    fields = (fragment["sha256"], fragment["sheet"], str(key & 0xffffffff),
              str(second), str(msg), "" if mmsi < 0 else str(mmsi), channel,
              format(bearing, ".17g"), format(distance, ".17g"))
    return hashlib.sha256("|".join(fields).encode("utf-8")).hexdigest()[:24]


def verified_fragments(cached, task):
    fragments = cached["fragment_info"]
    planned = {str(Path(path).resolve()).casefold() for path in task["paths"]}
    actual = [str(Path(fragment["path"]).resolve()).casefold() for fragment in fragments]
    if len(actual) != len(set(actual)) or set(actual) != planned:
        raise ValueError("fragment provenance paths differ from independently planned inputs")
    for fragment in fragments:
        if not isinstance(fragment.get("sheet"), str) or not fragment["sheet"]:
            raise ValueError("fragment sheet provenance is missing")
        digest = hashlib.sha256()
        with Path(fragment["path"]).open("rb") as handle:
            while chunk := handle.read(1024 * 1024):
                digest.update(chunk)
        if digest.hexdigest() != fragment.get("sha256"):
            raise ValueError("fragment provenance SHA differs from actual source bytes")
    if [fragment["sha256"] for fragment in fragments] != sorted(fragment["sha256"] for fragment in fragments):
        raise ValueError("fragment provenance order differs from canonical digest order")
    return fragments


def displayed_candidates(records, selection, fragments, day, limit=50):
    """Independent exact display contract, including deep selected cluster rows."""
    descending = sorted((record for record in records if record[0] <= 500.),
                        key=lambda record: (-record[0], record[1]))
    positions = list(range(min(limit, len(descending))))
    rank = selection["selected_rank"]
    if rank is not None and rank > limit:
        positions += list(range(rank - 1, min(len(descending), rank - 1 +
                                             min(limit, selection["cluster_count"]))))
    wanted = []
    for position in positions:
        record = descending[position]
        distance, key, bearing, second, msg, mmsi, channel = record
        fragment = fragments[key >> 32]
        wanted.append({"distance": distance, "bearing": bearing, "source_row": key & 0xffffffff,
                       "rank": position + 1, "candidate_id": candidate_identity(fragment, record),
                       "timestamp": (datetime.fromisoformat(day) + timedelta(seconds=second)).isoformat(),
                       "mmsi": None if mmsi < 0 else mmsi, "channel": channel,
                       "fragment": Path(fragment["path"]).name, "sheet": fragment["sheet"], "msg_type": msg})
    return wanted


def audit_task(task: dict[str, object]) -> dict[str, object]:
    started = time.perf_counter()
    before = input_stamps(task)
    payload = json.loads(Path(str(task["cache"])).read_text(encoding="utf-8"))
    signature, cached = payload["signature"], payload["result"]
    if (set(signature["message_types"]) != MODERN or signature["max_distance"] != 500
            or signature["tolerance"] != .10 or signature["cluster_size"] != 3
            or signature.get("top_candidates") != 50
            or signature["date"] != task["day"] or cached["day"] != task["day"]
            or signature["port"] != task["port"]):
        raise ValueError("cache does not match this audit's documented default-rule contract")
    scopes = {"logical_day": cached["directions"], **cached["period_directions"]}
    fragments = verified_fragments(cached, task)
    if set(scopes) != {"logical_day", "period_a", "period_b"}:
        raise ValueError("missing or unexpected selector scopes")
    if any(set(values) != set(LABELS) for values in scopes.values()):
        raise ValueError("cache lacks the complete documented 32 direction keys")
    issues: list[dict[str, object]] = []
    selections: list[dict[str, object]] = []
    digest = hashlib.sha256()
    row_count = 0
    candidate_rows = 0
    with Path(str(task["spool"])).open("rb") as handle:
        magic = handle.read(8)
        digest.update(magic)
        if magic != b"AISNRM1\0":
            raise ValueError("invalid normalized spool magic")
        header = handle.read(COUNTS.size)
        digest.update(header)
        counts = COUNTS.unpack(header)
        if Path(str(task["spool"])).stat().st_size != 8 + COUNTS.size + sum(counts) * ROW.size:
            raise ValueError("spool size differs from complete row counts")
        for index, count in enumerate(counts):
            buckets: dict[str, list[tuple[float, int, float]]] = {
                "logical_day": [], "period_a": [], "period_b": [],
            }
            remaining = count
            keys = set()
            while remaining:
                batch = min(remaining, 16384)
                data = handle.read(batch * ROW.size)
                digest.update(data)
                if len(data) != batch * ROW.size:
                    raise ValueError("truncated spool records")
                for distance, bearing, key, second, msg, period, mmsi, channel in ROW.iter_unpack(data):
                    if (second >= 86400 or period != int(second >= 43200)
                            or exact_float_bin(bearing) != index or mmsi < -1):
                        raise ValueError("spool direction/time/period is inconsistent")
                    fragment_index, source_row = key >> 32, key & 0xffffffff
                    if (key in keys or fragment_index >= len(fragments)
                            or not 2 <= source_row <= fragments[fragment_index]["rows_scanned"] + 1):
                        raise ValueError("spool source key is duplicate or outside fragment provenance")
                    keys.add(key)
                    record = (distance, key, bearing, second, msg, mmsi,
                              channel.rstrip(b"\0").decode("utf-8"))
                    if msg in MODERN:
                        buckets["logical_day"].append(record)
                    if msg in HISTORICAL:
                        buckets["period_b" if period else "period_a"].append(record)
                row_count += batch
                remaining -= batch
            for scope, records in buckets.items():
                expected = oracle(records)
                selected = next((row for row in records if row[1] == expected["selected_source_key"]), None)
                expected["selected_candidate_id"] = (None if selected is None else
                    candidate_identity(fragments[selected[1] >> 32], selected))
                actual = scopes[scope][LABELS[index]]
                candidates = displayed_candidates(records, expected, fragments, str(task["day"]))
                observed = actual.get("candidates")
                if not isinstance(observed, list) or len(observed) != len(candidates):
                    issues.append({"scope": scope, "index": index, "field": "candidates",
                                   "reason": "display candidate count/coverage differs"})
                else:
                    for position, (wanted, seen) in enumerate(zip(candidates, observed)):
                        candidate_rows += 1
                        if any(field not in seen or not equal_field(value, seen[field])
                               for field, value in wanted.items()):
                            issues.append({"scope": scope, "index": index, "field": "candidates",
                                           "position": position, "reason": "source/value/rank/identity metadata differs"})
                selections.append({"scope": scope, "index": index, **expected})
                for field in FIELDS:
                    if field not in actual or not equal_field(expected[field], actual.get(field)):
                        issues.append({"scope": scope, "index": index, "direction": LABELS[index],
                                       "field": field, "expected": expected[field], "actual": actual.get(field)})
        if handle.read(1):
            raise ValueError("unexpected trailing spool bytes")
    if digest.hexdigest() != str(cached["spool_sha256"]).lower():
        issues.append({"field": "spool_sha256", "reason": "cache/spool integrity differs"})
    if before != input_stamps(task):
        issues.append({"field": "input_stamps", "reason": "source/cache/spool changed during audit"})
    result = {"port": task["port"], "day": task["day"], "rows_scanned": row_count,
              "selectors_checked": len(selections), "numeric_fields_checked": len(selections) * 6,
              "identity_fields_checked": len(selections),
              "candidate_rows_checked": candidate_rows,
              "selections": selections, "issues": issues, "pass": not issues,
              "input_stamps": before, "elapsed_seconds": round(time.perf_counter() - started, 3)}
    return seal_entry(result, task)


def retained_results(previous, tasks, source_root, run_root):
    """Recheck sealed same-root evidence against exact live inputs and cache."""
    if (previous.get("schema") != SCHEMA
            or previous.get("source_root") != str(source_root.resolve())
            or previous.get("run_root") != str(run_root.resolve())
            or previous.get("planned_days") != len(tasks)):
        raise ValueError("selector checkpoint schema/root/plan differs")
    by_key = {(task["port"], task["day"]): task for task in tasks}
    entries = previous.get("results")
    if not isinstance(entries, list) or previous.get("completed_days") != len(entries):
        raise ValueError("selector checkpoint coverage is malformed")
    seen, retained = set(), []
    for entry in entries:
        key = entry.get("port"), entry.get("day")
        if key in seen or key not in by_key:
            raise ValueError("selector checkpoint duplicate or unplanned day")
        seen.add(key)
        if entry.get("pass") is False:
            continue  # Failed/partial tasks are re-audited, never retained as PASS.
        body = {name: value for name, value in entry.items() if name != "entry_sha256"}
        encoded = json.dumps(body, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        if (entry.get("pass") is not True or entry.get("issues") != []
                or entry.get("entry_sha256") != hashlib.sha256(encoded.encode("utf-8")).hexdigest()
                or entry.get("input_stamps") != input_stamps(by_key[key])):
            raise ValueError("selector checkpoint PASS seal/input stamps differ")
        payload = json.loads(Path(str(by_key[key]["cache"])).read_text(encoding="utf-8"))["result"]
        scopes = {"logical_day": payload["directions"], **payload["period_directions"]}
        selections = entry.get("selections")
        wanted = {(scope, index) for scope in scopes for index in range(32)}
        observed = [(selection.get("scope"), selection.get("index")) for selection in selections or []]
        if (len(observed) != 96 or len(set(observed)) != 96 or set(observed) != wanted
                or entry.get("selectors_checked") != 96 or entry.get("numeric_fields_checked") != 576
                or entry.get("identity_fields_checked") != 96
                or entry.get("candidate_rows_checked") != sum(len(item["candidates"])
                    for directions in scopes.values() for item in directions.values())
                or entry.get("rows_scanned") != payload["rows_normalized"]):
            raise ValueError("selector checkpoint row/field/32-direction coverage differs")
        for selection in selections:
            current = scopes[selection["scope"]][LABELS[selection["index"]]]
            if any(field not in current or not equal_field(selection.get(field), current[field]) for field in FIELDS):
                raise ValueError("selector checkpoint selections differ from current cache")
        retained.append(entry)
    return retained


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source_root", type=Path)
    parser.add_argument("run_root", type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--workers", type=int, default=2)
    parser.add_argument("--resume", action="store_true")
    args = parser.parse_args(argv)
    if not 1 <= args.workers <= 8:
        parser.error("workers must be 1..8")
    tasks = plan_tasks(args.source_root, args.run_root)
    if not tasks:
        raise ValueError("empty source cannot pass")
    inputs = [Path(str(task[key])) for task in tasks for key in ("cache", "spool")]
    inputs += [Path(path) for task in tasks for path in task["paths"]]
    protected = [args.source_root, args.run_root]
    results: list[dict[str, object]] = []
    if args.output.exists():
        if not args.resume:
            raise ValueError("existing selector evidence requires explicit --resume")
        previous = json.loads(args.output.read_text(encoding="utf-8"))
        results = retained_results(previous, tasks, args.source_root, args.run_root)
    report = {"schema": SCHEMA, "source_root": str(args.source_root.resolve()),
              "run_root": str(args.run_root.resolve()), "planned_days": len(tasks), "completed_days": 0,
              "boundary": "ALL 32 cache directions and all 3 scopes, ALL spool records; six numerical fields and selected provenance ID. "
                          "ALL displayed candidates (including deep selected clusters) independently bound to source provenance. "
                          "Requires separate independent raw/spool census INCLUDING source-key provenance and final workbook/visual verification.",
              "pass": False, "results": results}
    report["completed_days"] = len(results)
    report["resumed_days"] = len(results)
    report["pass"] = len(results) == len(tasks) and all(item["pass"] for item in results)
    save_report(args.output, report, protected, inputs)
    retained = {(entry["port"], entry["day"]) for entry in results}
    with ProcessPoolExecutor(max_workers=args.workers) as pool:
        futures = {pool.submit(audit_task, task): task for task in tasks
                   if (task["port"], task["day"]) not in retained}
        for future in as_completed(futures):
            task = futures[future]
            try:
                result = future.result()
            except Exception as error:
                result = {"port": task["port"], "day": task["day"], "pass": False,
                          "issues": [{"error": f"{type(error).__name__}: {error}"}]}
            results.append(result)
            results.sort(key=lambda item: (item["port"], item["day"]))
            report["completed_days"] = len(results)
            report["pass"] = len(results) == len(tasks) and all(item["pass"] for item in results)
            save_report(args.output, report, protected, inputs)
            print(f"{len(results)}/{len(tasks)} {task['port']} {task['day']} {'PASS' if result['pass'] else 'FAIL'}", flush=True)
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
