"""Independent raw-to-cache/spool census; no production imports or selector oracle.

Raw XML numbers are parsed as finite binary64, as Excel readers do. Direction
bins are evaluated as an exact integer ratio of that binary64 number, not by
calling the application's parser/binning functions. The mathematical rule is
floor((bearing modulo 360) * 4 / 45). Raw/spool records are compared as multisets
using counts, period/message histograms and commutative SHA256 accumulators.
The accumulators bind every value to its independently reconstructed source key:
SHA256-sorted source fragment index and actual XML/Excel row number. Cache fragment
hashes, paths and worksheet titles must match the independently read originals.
This proves normalization coverage under the documented rules, not manual
selector equivalence or chart fidelity. All workbook access is read-only.
"""
from __future__ import annotations

import argparse
from collections import Counter
from concurrent.futures import ProcessPoolExecutor, as_completed
from datetime import date, datetime
import hashlib
import json
import math
from pathlib import Path
import posixpath
import re
import struct
import time
from typing import Callable
from xml.parsers import expat
import xml.etree.ElementTree as ET
import zipfile

SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
HISTORICAL_TYPES = frozenset((1, 3, 4, 18, 19))
FIELDS = {
    "year": ("year",), "month": ("month",), "day": ("day",),
    "hour": ("hour",), "minute": ("minute",), "second": ("second",),
    "channel": ("channel",), "msg_type": ("msg_type", "msg type"),
    "mmsi": ("mmsi",), "longitude_desc": ("longitude_desc", "longitude desc"),
    "bearing": ("bearing",),
    "distance": ("distance in nautical miles", "distance_in_nautical_miles"),
}
# Evidence-only names used by the pre-v1.5.1 June manual workbooks.
REFERENCE_LABELS = (
    "北", "北微東", "北北東", "東北微北", "東北", "東北微東", "東北東", "東微北",
    "東", "東微南", "東南東", "東南微南", "東南", "東南微西", "南南東", "南微東",
    "南", "南微西", "南南西", "西南微南", "西南", "西南微西", "西南西", "西微南",
    "西", "西微北", "西北西", "西北微西", "西北", "西北微北", "北北西", "北微西",
)
COUNTER_NAMES = (
    "rows_scanned", "rows_accepted", "rows_invalid", "rows_wrong_message",
    "rows_not_east", "rows_legacy", "rows_normalized",
)
SPOOL_RECORD = struct.Struct("<ddQIIBq16s")
IDENTITY_RECORD = struct.Struct("<ddQIIq16s")
COUNTS = struct.Struct("<32Q")
HASH_MOD = 1 << 256
CHECKPOINT_SCHEMA = "AIS_INDEPENDENT_SOURCE_AUDIT_3"
SOURCE_IDENTITY_RULE = "sha256-sorted-fragment-index<<32 | actual-excel-row; bound to every normalized value"


def independent_identity(name: str) -> tuple[str, date]:
    match = re.fullmatch(
        r"D&TMOK[ \t]+([A-Z][A-Z0-9]{1,15})_(\d{7,8})(?:_[A-Z0-9_-]+)?\.xlsx",
        name, re.IGNORECASE,
    )
    if match is None:
        raise ValueError(f"unrecognized AIS filename: {name}")
    digits = match[2]
    return match[1].upper(), date(int(digits[:4]), int(digits[4:6]), int(digits[6:]))


def finite_number(value: object) -> float | None:
    try:
        number = float(value)  # cached XML numeric value is read as binary64
    except (TypeError, ValueError, OverflowError):
        return None
    return number if math.isfinite(number) else None


def exact_float_bin(value: object) -> int | None:
    number = finite_number(value)
    if number is None:
        return None
    numerator, denominator = number.as_integer_ratio()
    remainder = numerator % (360 * denominator)
    return (remainder * 4) // (45 * denominator)


def integer(value: object) -> int | None:
    number = finite_number(value)
    return int(number) if number is not None and number.is_integer() else None


def column_number(reference: str) -> int:
    value = 0
    for char in reference:
        if not char.isalpha():
            break
        value = value * 26 + ord(char.upper()) - ord("A") + 1
    if not value:
        raise ValueError(f"cell has no column reference: {reference!r}")
    return value - 1


def read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    strings: list[str] = []
    with archive.open("xl/sharedStrings.xml") as handle:
        for _, element in ET.iterparse(handle, events=("end",)):
            if element.tag == f"{{{SHEET_NS}}}si":
                strings.append("".join(node.text or "" for node in element.iter(f"{{{SHEET_NS}}}t")))
                element.clear()
    return strings


def workbook_sheets(archive: zipfile.ZipFile) -> list[tuple[str, str]]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        item.attrib["Id"]: item.attrib["Target"]
        for item in relationships
        if item.attrib.get("TargetMode") != "External"
    }
    sheets = []
    for item in workbook.findall(f"{{{SHEET_NS}}}sheets/{{{SHEET_NS}}}sheet"):
        target = targets[item.attrib[f"{{{REL_NS}}}id"]]
        part = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
        sheets.append((item.attrib["name"], part))
    return sheets


def stream_sheet(
    archive: zipfile.ZipFile, part: str, shared: list[str],
    callback: Callable[[int, dict[int, object]], set[int] | None],
    columns: set[int] | None = None,
) -> None:
    """SAX worksheet reader. Retains one row, only requested cells after header."""
    parser = expat.ParserCreate(namespace_separator="}")
    parser.buffer_text = True
    row_index = 0
    row: dict[int, object] = {}
    column = -1
    cell_type = ""
    buffer: list[str] = []
    in_value = False
    selected = columns

    def start(name: str, attrs: dict[str, str]) -> None:
        nonlocal row_index, row, column, cell_type, buffer, in_value
        tag = name.rsplit("}", 1)[-1]
        if tag == "row":
            row_index = int(attrs.get("r", row_index + 1))
            row = {}
        elif tag == "c":
            column = column_number(attrs["r"])
            cell_type = attrs.get("t", "n")
            buffer = []
        elif tag in ("v", "t"):
            in_value = selected is None or column in selected

    def chars(text: str) -> None:
        if in_value:
            buffer.append(text)

    def end(name: str) -> None:
        nonlocal in_value, selected
        tag = name.rsplit("}", 1)[-1]
        if tag in ("v", "t"):
            in_value = False
        elif tag == "c" and (selected is None or column in selected):
            text = "".join(buffer)
            if cell_type == "s":
                row[column] = shared[int(text)] if text else None
            elif cell_type in ("inlineStr", "str"):
                row[column] = text
            elif cell_type == "e":
                row[column] = None
            else:
                row[column] = text if text else None
        elif tag == "row":
            next_columns = callback(row_index, row)
            if next_columns is not None:
                selected = next_columns

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.CharacterDataHandler = chars
    with archive.open(part) as handle:
        while chunk := handle.read(1024 * 1024):
            parser.Parse(chunk, False)
        parser.Parse(b"", True)


def identify_columns(header: dict[int, object]) -> dict[str, int]:
    normalized = {
        index: " ".join(str(value or "").strip().casefold().split())
        for index, value in header.items()
    }
    indexes = {}
    for field, aliases in FIELDS.items():
        matches = [index for index, name in normalized.items() if name in aliases]
        if len(matches) != 1:
            raise ValueError(f"{field}: expected one column, found {matches}")
        indexes[field] = matches[0]
    return indexes


class HeaderRead(Exception):
    pass


def locate_sheet(archive: zipfile.ZipFile, shared: list[str]) -> tuple[str, str, dict[str, int]]:
    candidates = []
    for title, part in workbook_sheets(archive):
        header: dict[int, object] = {}

        def first(row_index: int, row: dict[int, object]) -> None:
            if row_index != 1:
                raise ValueError(f"{title}: header is not row 1")
            header.update(row)
            raise HeaderRead

        try:
            stream_sheet(archive, part, shared, first)
        except HeaderRead:
            pass
        try:
            candidates.append((title, part, identify_columns(header)))
        except ValueError:
            continue
    ais = [candidate for candidate in candidates if candidate[0].strip().casefold() == "ais"]
    if len(ais) == 1:
        return ais[0]
    if len(candidates) != 1:
        raise ValueError(f"ambiguous/missing data sheets: {[item[0] for item in candidates]}")
    return candidates[0]


class RecordCensus:
    def __init__(self) -> None:
        self.counts = [0] * 32
        self.hash_sum = [0] * 32
        self.hash_xor = [0] * 32
        self.histogram: Counter[tuple[int, int, int]] = Counter()

    def add(self, index: int, distance: float, bearing: float, second: int,
            msg_type: int, mmsi: int | None, channel: str, *, source_key: int) -> None:
        if not isinstance(source_key, int) or not 0 <= source_key < 1 << 64 or (source_key & 0xFFFFFFFF) < 2:
            raise ValueError("invalid source key / data row must be at least 2")
        if mmsi is not None and mmsi < 0:
            raise ValueError("negative source MMSI conflicts with spool missing-value sentinel")
        encoded = channel.encode("utf-8")
        if len(encoded) > 16:
            raise ValueError("channel exceeds normalized spool's 16 UTF-8 bytes")
        identity = IDENTITY_RECORD.pack(
            distance, bearing, source_key, second, msg_type, -1 if mmsi is None else mmsi,
            encoded.ljust(16, b"\0"),
        )
        digest = int.from_bytes(hashlib.sha256(identity).digest(), "big")
        self.counts[index] += 1
        self.hash_sum[index] = (self.hash_sum[index] + digest) % HASH_MOD
        self.hash_xor[index] ^= digest
        self.histogram[index, int(second >= 43200), msg_type] += 1

    def evidence(self) -> dict[str, object]:
        return {
            "source_identity_rule": SOURCE_IDENTITY_RULE,
            "direction_counts": self.counts,
            "sha256_multiset_sum": [f"{value:064x}" for value in self.hash_sum],
            "sha256_multiset_xor": [f"{value:064x}" for value in self.hash_xor],
            "period_message_direction_groups": len(self.histogram),
        }


def scan_source(path: Path, expected_day: date, modern_types: set[int],
                max_distance: float, census: RecordCensus, fragment_index: int = 0) -> dict[str, object]:
    counters = {name: 0 for name in COUNTER_NAMES}
    occupied = bytearray(86400)
    min_second: int | None = None
    max_second: int | None = None
    last_row = 1
    boundaries = 0
    examples: list[dict[str, object]] = []
    with zipfile.ZipFile(path) as archive:
        shared = read_shared_strings(archive)
        title, part, indexes = locate_sheet(archive, shared)

        def process(row_index: int, row: dict[int, object]) -> set[int] | None:
            nonlocal last_row, min_second, max_second, boundaries
            if row_index == 1:
                return set(indexes.values())
            if row_index != last_row + 1:
                raise ValueError(f"{path.name}: missing/noncontiguous XML row {last_row + 1}")
            last_row = row_index
            counters["rows_scanned"] += 1
            components = [integer(row.get(indexes[key])) for key in (
                "year", "month", "day", "hour", "minute", "second",
            )]
            if any(value is None for value in components):
                raise ValueError(f"{path.name}:{row_index}: invalid date/time integer")
            timestamp = datetime(*components)  # type: ignore[arg-type]
            if timestamp.date() != expected_day:
                raise ValueError(f"{path.name}:{row_index}: {timestamp.date()} != {expected_day}")
            second = timestamp.hour * 3600 + timestamp.minute * 60 + timestamp.second
            occupied[second] = 1
            min_second = second if min_second is None else min(min_second, second)
            max_second = second if max_second is None else max(max_second, second)
            msg = integer(row.get(indexes["msg_type"]))
            if msg is None or not 0 <= msg <= 0xFFFFFFFF:
                counters["rows_invalid"] += 1
                return None
            if msg not in modern_types:
                counters["rows_wrong_message"] += 1
            if str(row.get(indexes["longitude_desc"])).strip().casefold() != "east":
                counters["rows_not_east"] += 1
                return None
            raw_bearing = row.get(indexes["bearing"])
            bearing = finite_number(raw_bearing)
            distance = finite_number(row.get(indexes["distance"]))
            index = exact_float_bin(bearing)
            if index is None or distance is None or distance < 0:
                counters["rows_invalid"] += 1
                return None
            normalized = bearing % 360.0  # type: ignore[operator]
            # Binary64 modulo can round a negative near-zero bearing to 360.
            # The independent exact-ratio bin remains authoritative; surface
            # any classification discrepancy instead of silently adapting it.
            float_index = min(math.floor(normalized * 4.0 / 45.0), 31)
            if index != float_index:
                boundaries += 1
                if len(examples) < 10:
                    examples.append({"row": row_index, "raw": raw_bearing,
                                     "exact_bin": index, "float_bin": float_index})
            mmsi = integer(row.get(indexes["mmsi"]))
            channel = "" if row.get(indexes["channel"]) is None else str(row[indexes["channel"]]).strip()
            if not 0 <= fragment_index < 1 << 32 or not 2 <= row_index < 1 << 32:
                raise ValueError("fragment index / Excel row exceeds source-key representation")
            census.add(index, distance, normalized, second, msg, mmsi, channel,
                       source_key=(fragment_index << 32) | row_index)
            counters["rows_normalized"] += 1
            counters["rows_legacy"] += int(msg in HISTORICAL_TYPES)
            counters["rows_accepted"] += int(msg in modern_types and distance <= max_distance)
            return None

        stream_sheet(archive, part, shared, process)
    return {
        "path": str(path), "sheet": title, "columns_zero_based": indexes,
        **counters, "min_second": min_second, "max_second": max_second,
        "occupied_seconds": sum(occupied), "binary_boundary_differences": boundaries,
        "boundary_examples": examples,
    }


def scan_spool(path: Path) -> RecordCensus:
    census = RecordCensus()
    with path.open("rb") as handle:
        if handle.read(8) != b"AISNRM1\0":
            raise ValueError(f"invalid spool magic: {path}")
        count_bytes = handle.read(COUNTS.size)
        if len(count_bytes) != COUNTS.size:
            raise ValueError("truncated spool counts")
        counts = COUNTS.unpack(count_bytes)
        if path.stat().st_size != 8 + COUNTS.size + sum(counts) * SPOOL_RECORD.size:
            raise ValueError("spool size does not match row counts")
        for index, count in enumerate(counts):
            remaining = count
            while remaining:
                batch_count = min(remaining, 16384)
                data = handle.read(batch_count * SPOOL_RECORD.size)
                if len(data) != batch_count * SPOOL_RECORD.size:
                    raise ValueError("truncated spool records")
                for distance, bearing, source_key, second, msg, period, mmsi, channel in SPOOL_RECORD.iter_unpack(data):
                    if period != int(second >= 43200):
                        raise ValueError("spool period code disagrees with timestamp")
                    if exact_float_bin(bearing) != index:
                        raise ValueError(f"spool direction bin {index} disagrees with bearing {bearing}")
                    census.add(index, distance, bearing, second, msg,
                               None if mmsi == -1 else mmsi, channel.rstrip(b"\0").decode("utf-8"),
                               source_key=source_key)
                remaining -= batch_count
    return census


def source_fingerprints(paths: list[Path]) -> list[dict[str, str]]:
    """Independently hash local originals and derive canonical fragment order."""
    result = []
    for path in paths:
        before = path.stat()
        digest = hashlib.sha256()
        with path.open("rb") as handle:
            while chunk := handle.read(1024 * 1024):
                digest.update(chunk)
        after = path.stat()
        fields = ("st_size", "st_mtime_ns", "st_dev", "st_ino")
        if any(getattr(before, name) != getattr(after, name) for name in fields):
            raise ValueError(f"source changed during fingerprint: {path}")
        result.append({"path": str(path.resolve()), "sha256": digest.hexdigest().upper()})
    if len({item["sha256"] for item in result}) != len(result):
        raise ValueError("byte-identical fragment collapse requires separately verified alias handling")
    return sorted(result, key=lambda item: item["sha256"])


def audit_day(task: dict[str, object]) -> dict[str, object]:
    started = time.perf_counter()
    cache_path = Path(str(task["cache"]))
    payload = json.loads(cache_path.read_text(encoding="utf-8"))
    signature, expected = payload["signature"], payload["result"]
    day = date.fromisoformat(str(task["day"]))
    issues: list[str] = []
    census = RecordCensus()
    fragments = []
    try:
        if signature["date"] != day.isoformat() or expected["day"] != day.isoformat():
            raise ValueError("cache signature/result day differs from independently parsed source")
        if str(signature["port"]).upper() != task["port"]:
            raise ValueError("cache port differs from independently parsed source")
        if (set(signature["message_types"]) != {1, 2, 3, 18, 19}
                or float(signature["max_distance"]) != 500.0):
            raise ValueError("cache differs from this regression's documented Modern profile / 500 NM cap")
        expected_policy = "klng21" if task["port"] == "KLNG" else "all32"
        if signature.get("direction_policy") != expected_policy:
            raise ValueError("cache direction policy differs from independently applied port contract")
        paths = [Path(path) for path in task["paths"]]  # type: ignore[union-attr]
        cached_paths = sorted(str(Path(info["path"]).resolve()) for info in signature["sources"])
        if cached_paths != sorted(str(path.resolve()) for path in paths):
            raise ValueError("cache source list differs from independently inventoried input")
        fingerprints = source_fingerprints(paths)
        expected_fragments = expected["fragment_info"]
        if len(expected_fragments) != len(paths):
            raise ValueError("canonical/byte-identical fragment collapse requires separate investigation")
        for fragment_index, fingerprint in enumerate(fingerprints):
            path = Path(fingerprint["path"])
            expected_fragment = expected_fragments[fragment_index]
            if (str(Path(expected_fragment["path"]).resolve()) != fingerprint["path"]
                    or str(expected_fragment.get("sha256", "")).upper() != fingerprint["sha256"]
                    or expected_fragment.get("aliases", [])):
                raise ValueError("cache canonical fragment path/hash/index differs from independently hashed source")
            fragment = scan_source(path, day, set(signature["message_types"]),
                                   float(signature["max_distance"]), census, fragment_index)
            fragment["source_sha256"] = fingerprint["sha256"]
            fragment["fragment_index"] = fragment_index
            fragments.append(fragment)
            for metric in ("sheet", "rows_scanned", "min_second", "max_second", "occupied_seconds"):
                if fragment[metric] != expected_fragment[metric]:
                    issues.append(f"{path.name}:{metric}: raw={fragment[metric]} cache={expected_fragment[metric]}")
            if fragment["binary_boundary_differences"]:
                issues.append(f"{path.name}: exact/float bin boundary discrepancy")
        totals = {name: sum(int(fragment[name]) for fragment in fragments) for name in COUNTER_NAMES}
        for metric, value in totals.items():
            if value != expected[metric]:
                issues.append(f"{metric}: raw={value} cache={expected[metric]}")
        spool = scan_spool(Path(str(task["spool"])))
        if census.histogram != spool.histogram:
            issues.append("raw/spool direction-period-message count mismatch")
        if census.evidence() != spool.evidence():
            issues.append("raw/spool normalized record multiset/count mismatch")
        return {
            "port": task["port"], "day": task["day"], "cache": str(cache_path),
            "spool": task["spool"], "source_files": len(paths), "fragments": fragments,
            "source_fingerprints": fingerprints,
            "totals": totals, "raw_census": census.evidence(), "spool_census": spool.evidence(),
            "issues": issues, "pass": not issues,
            "elapsed_seconds": round(time.perf_counter() - started, 3),
        }
    except Exception as error:
        return {"port": task["port"], "day": task["day"], "fragments": fragments,
                "issues": [*issues, f"{type(error).__name__}: {error}"], "pass": False,
                "elapsed_seconds": round(time.perf_counter() - started, 3)}


def plan_tasks(source_root: Path, run_root: Path) -> list[dict[str, object]]:
    sources: dict[tuple[str, date], list[str]] = {}
    for path in sorted(source_root.rglob("*.xlsx")):
        port, day = independent_identity(path.name)
        sources.setdefault((port, day), []).append(str(path))
    caches: dict[tuple[str, date], Path] = {}
    for path in run_root.rglob("*.json"):
        if not re.fullmatch(r"\d{8}\.json", path.name) or not path.parent.name.endswith("_cache"):
            continue
        payload = json.loads(path.read_text(encoding="utf-8"))
        signature = payload["signature"]
        key = str(signature["port"]), date.fromisoformat(signature["date"])
        if key in caches:
            raise ValueError(f"multiple caches for {key}")
        caches[key] = path
    if set(sources) != set(caches):
        raise ValueError(f"source/cache days differ: missing={set(sources)-set(caches)} extra={set(caches)-set(sources)}")
    tasks = []
    for (port, day), paths in sorted(sources.items()):
        cache = caches[port, day]
        spool_dir = cache.parent.with_name(cache.parent.name.removesuffix("_cache") + "_rows")
        spool = spool_dir / (day.strftime("%Y%m%d") + ".bin")
        if not spool.is_file():
            raise ValueError(f"missing spool: {spool}")
        tasks.append({"port": port, "day": day.isoformat(), "paths": paths,
                      "cache": str(cache), "spool": str(spool)})
    return tasks


def audit_reference(path: Path) -> dict[str, object]:
    started = time.perf_counter()
    counts: Counter[str] = Counter()
    sheet_counts: dict[str, int] = {}
    examples: list[dict[str, object]] = []
    labels = {label: index for index, label in enumerate(REFERENCE_LABELS)}
    with zipfile.ZipFile(path) as archive:
        shared = read_shared_strings(archive)
        for title, part in workbook_sheets(archive):
            if title == "總表":
                continue
            sheet_counts[title] = 0

            def inspect(row_index: int, row: dict[int, object]) -> None:
                if row_index == 1:
                    return
                distance, bearing, label = row.get(0), row.get(1), row.get(2)
                if distance is None and bearing is None and label is None:
                    counts["blank_rows"] += 1
                    return
                counts["detail_rows"] += 1
                sheet_counts[title] += 1
                expected_index = labels.get(str(label))
                actual = exact_float_bin(bearing)
                if expected_index is None or actual is None:
                    counts["unrecognized_rows"] += 1
                    reason = "unknown legacy direction label or nonfinite bearing"
                elif actual != expected_index:
                    counts["mapping_mismatches"] += 1
                    reason = "bearing bin differs from legacy label"
                else:
                    counts["matched_rows"] += 1
                    return
                if len(examples) < 20:
                    examples.append({"sheet": title, "row": row_index, "bearing": bearing,
                                     "label": label, "bin": actual, "reason": reason})

            stream_sheet(archive, part, shared, inspect, {0, 1, 2})
    return {"path": str(path), "coverage": "ALL detail sheets; no row cap",
            **{key: counts[key] for key in ("detail_rows", "matched_rows", "blank_rows",
                                          "unrecognized_rows", "mapping_mismatches")},
            "sheet_rows": sheet_counts, "examples": examples,
            "pass": counts["detail_rows"] > 0 and counts["mapping_mismatches"] == 0
                    and counts["unrecognized_rows"] == 0,
            "elapsed_seconds": round(time.perf_counter() - started, 3)}


def guard_output(path: Path, protected_roots: list[Path], protected_files: list[Path]) -> None:
    """Check destination and atomic temporary before any mkdir/write/replace.

    Resolve symlinks/junctions and compare hard-link identity with every input.
    The caller owns its report directory; source/cache/reference trees remain
    read-only. Checks do not defend against a concurrent hostile path swap.
    """
    if path.suffix.casefold() != ".json":
        raise ValueError("audit output must have .json extension")
    temporary = path.with_suffix(path.suffix + ".tmp")
    roots = [root.resolve() for root in protected_roots]
    inputs = [item.resolve() for item in protected_files]
    for candidate in (path, temporary):
        resolved = candidate.resolve()
        if any(resolved == root or root in resolved.parents for root in roots):
            raise ValueError(f"report/temporary lies within a protected input tree: {candidate}")
        if any(resolved == item for item in inputs):
            raise ValueError(f"report/temporary aliases a protected input: {candidate}")
        if candidate.exists():
            if not candidate.is_file():
                raise ValueError(f"report/temporary is not a regular file: {candidate}")
            for item in protected_files:
                if item.exists() and candidate.samefile(item):
                    raise ValueError(f"report/temporary hard-links a protected input: {candidate}")


def input_stamps(task: dict[str, object]) -> list[dict[str, object]]:
    paths = [Path(path) for path in task["paths"]]  # type: ignore[union-attr]
    paths.extend((Path(str(task["cache"])), Path(str(task["spool"]))))
    result = []
    for path in sorted(paths, key=lambda item: str(item.resolve()).casefold()):
        info = path.stat()
        result.append({"path": str(path.resolve()), "size": info.st_size,
                       "modified_ns": info.st_mtime_ns})
    return result


def seal_entry(result: dict[str, object], task: dict[str, object]) -> dict[str, object]:
    sealed = dict(result)
    if "input_stamps" not in sealed:
        sealed["input_stamps"] = input_stamps(task)
    encoded = json.dumps(sealed, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    sealed["entry_sha256"] = hashlib.sha256(encoded.encode("utf-8")).hexdigest()
    return sealed


def audit_task_with_stamps(task: dict[str, object]) -> dict[str, object]:
    before = input_stamps(task)
    result = audit_day(task)
    if before != input_stamps(task):
        result["issues"].append("source/cache/spool changed during the independent audit")  # type: ignore[union-attr]
        result["pass"] = False
    result["input_stamps"] = before
    return seal_entry(result, task)


def resume_results(previous: dict[str, object], tasks: list[dict[str, object]],
                   source_root: Path, run_root: Path) -> list[dict[str, object]]:
    """Validate checkpoints before reusing success; failures are always retried.

    Entry hashes detect accidental edits such as changing FAIL to PASS. They
    are not signed authentication; re-hashing a malicious edit is out of scope.
    File stat stamps detect changed task inputs, not same-size/mtime tampering.
    The separate source SHA audit establishes stronger immutable provenance.
    """
    if previous.get("checkpoint_schema") != CHECKPOINT_SCHEMA:
        raise ValueError("unsupported/unsealed resume report; do not trust old PASS entries")
    if (previous.get("source_root") != str(source_root.resolve())
            or previous.get("run_root") != str(run_root.resolve())):
        raise ValueError("resume roots differ from prior report")
    by_key = {(task["port"], task["day"]): task for task in tasks}
    seen: set[tuple[object, object]] = set()
    retained = []
    entries = previous.get("results")
    if not isinstance(entries, list):
        raise ValueError("resume results are not a list")
    for entry in entries:
        if not isinstance(entry, dict):
            raise ValueError("malformed resume day entry")
        key = entry.get("port"), entry.get("day")
        if key not in by_key or key in seen:
            raise ValueError(f"extra/duplicate resume day: {key}")
        seen.add(key)
        checksum = entry.get("entry_sha256")
        body = {name: value for name, value in entry.items() if name != "entry_sha256"}
        encoded = json.dumps(body, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        if checksum != hashlib.sha256(encoded.encode("utf-8")).hexdigest():
            raise ValueError(f"tampered/unsealed resume entry: {key}")
        task = by_key[key]
        if entry.get("input_stamps") != input_stamps(task):
            raise ValueError(f"resume task inputs changed: {key}")
        if not isinstance(entry.get("issues"), list) or not isinstance(entry.get("pass"), bool):
            raise ValueError(f"malformed resume issues/pass: {key}")
        if not entry["pass"]:
            continue
        if entry["issues"]:
            raise ValueError(f"PASS resume entry retains issues: {key}")
        if entry.get("cache") != str(task["cache"]) or entry.get("spool") != str(task["spool"]):
            raise ValueError(f"resume cache/spool outside planned task: {key}")
        if entry.get("source_files") != len(task["paths"]):  # type: ignore[arg-type]
            raise ValueError(f"resume source count differs from planned task: {key}")
        raw_census, spool_census = entry.get("raw_census"), entry.get("spool_census")
        if not isinstance(raw_census, dict) or raw_census != spool_census:
            raise ValueError(f"resume has missing/different raw/spool census: {key}")
        if raw_census.get("source_identity_rule") != SOURCE_IDENTITY_RULE:
            raise ValueError(f"resume lacks source-key/value provenance census: {key}")
        direction_counts = raw_census.get("direction_counts")
        if (not isinstance(direction_counts, list) or len(direction_counts) != 32
                or any(not isinstance(value, int) or value < 0 for value in direction_counts)):
            raise ValueError(f"resume has invalid direction counts: {key}")
        totals = entry.get("totals")
        expected = json.loads(Path(str(task["cache"])).read_text(encoding="utf-8"))["result"]
        if (not isinstance(totals, dict)
                or any(totals.get(metric) != expected[metric] for metric in COUNTER_NAMES)
                or sum(direction_counts) != totals.get("rows_normalized")):
            raise ValueError(f"resume census counters differ from current cache: {key}")
        retained.append(entry)
    return retained


def save_report(path: Path, payload: dict[str, object], protected_roots: list[Path],
                protected_files: list[Path]) -> None:
    guard_output(path, protected_roots, protected_files)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temporary.replace(path)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="mode", required=True)
    raw = subparsers.add_parser("raw")
    raw.add_argument("source_root", type=Path)
    raw.add_argument("run_root", type=Path)
    raw.add_argument("--output", required=True, type=Path)
    raw.add_argument("--workers", type=int, default=2)
    raw.add_argument("--resume", action="store_true")
    raw.add_argument("--plan-only", action="store_true")
    reference = subparsers.add_parser("reference")
    reference.add_argument("workbooks", nargs="+", type=Path)
    reference.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    if args.mode == "reference":
        protected_files = [path.resolve() for path in args.workbooks]
        protected_roots = list({path.parent for path in protected_files})
        guard_output(args.output, protected_roots, protected_files)
        results = []
        for path in args.workbooks:
            results.append(audit_reference(path))
            save_report(args.output, {"rule": "exact binary64 rational modulo-360 floor-bin",
                                     "results": results, "pass": all(item["pass"] for item in results)},
                        protected_roots, protected_files)
            print(json.dumps(results[-1], ensure_ascii=False), flush=True)
        return 0 if all(item["pass"] for item in results) else 1
    if args.workers < 1 or args.workers > 8:
        parser.error("workers must be between 1 and 8")
    if str(args.source_root).startswith(("\\\\", "//")):
        parser.error("raw audit must use an isolated local source copy, not a network/UNC source")
    args.source_root = args.source_root.resolve()
    args.run_root = args.run_root.resolve()
    protected_roots = [args.source_root, args.run_root]
    guard_output(args.output, protected_roots, [])
    tasks = plan_tasks(args.source_root, args.run_root)
    protected_files = [Path(path) for task in tasks for path in task["paths"]]  # type: ignore[union-attr]
    protected_files.extend(Path(str(task[key])) for task in tasks for key in ("cache", "spool"))
    guard_output(args.output, protected_roots, protected_files)
    if args.plan_only:
        print(json.dumps({"source_files": sum(len(task["paths"]) for task in tasks),
                          "logical_days": len(tasks), "port_months": sorted({str(task['port']) + '_' + str(task['day'])[:7] for task in tasks})}, indent=2))
        return 0
    results = []
    if args.resume and args.output.is_file():
        previous = json.loads(args.output.read_text(encoding="utf-8"))
        results = resume_results(previous, tasks, args.source_root, args.run_root)
    completed = {(item["port"], item["day"]) for item in results}
    if not completed.issubset({(task["port"], task["day"]) for task in tasks}):
        raise ValueError("resume report has extra days")
    report: dict[str, object] = {
        "checkpoint_schema": CHECKPOINT_SCHEMA,
        "source_root": str(args.source_root), "run_root": str(args.run_root),
        "expected_source_files": sum(len(task["paths"]) for task in tasks),
        "expected_logical_days": len(tasks), "results": results,
        "independence": "stdlib OOXML SAX; independent filename/header/number/time/bin parser; no production imports",
        "scope": "all raw rows vs cache counters and normalized spool record multisets; selector/manual judgments/charts excluded",
        "resume_boundary": "sealed day results + same planned paths/stat stamps; checksum detects accidental edits, not signed authentication; separate immutable SHA audit required",
    }
    with ProcessPoolExecutor(max_workers=args.workers) as executor:
        futures = {executor.submit(audit_task_with_stamps, task): task for task in tasks
                   if (task["port"], task["day"]) not in completed}
        for future in as_completed(futures):
            result = future.result()
            results.append(result)
            results.sort(key=lambda item: (str(item["port"]), str(item["day"])))
            report["completed_days"] = len(results)
            report["pass"] = len(results) == len(tasks) and all(item["pass"] for item in results)
            save_report(args.output, report, protected_roots, protected_files)
            print(f"{len(results)}/{len(tasks)} {result['port']} {result['day']} "
                  f"{'PASS' if result['pass'] else 'FAIL'} {result['elapsed_seconds']}s", flush=True)
    report["completed_days"] = len(results)
    report["pass"] = len(results) == len(tasks) and all(item["pass"] for item in results)
    save_report(args.output, report, protected_roots, protected_files)
    return 0 if report.get("pass") else 1


if __name__ == "__main__":
    raise SystemExit(main())
