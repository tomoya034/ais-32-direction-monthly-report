from __future__ import annotations

import argparse
import calendar
import concurrent.futures
import hashlib
import json
import math
import multiprocessing
import os
import queue
import re
import shutil
import statistics
import struct
import subprocess
import sys
import threading
import time
import traceback
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Callable, Iterable, Sequence

import openpyxl
import xlsxwriter


APP_TITLE = "AIS 32方位月報一鍵製作"
APP_VERSION = "1.4.0"
CACHE_VERSION = 4
NORMALIZED_SPOOL_VERSION = 1
EXCEL_MAX_DATA_ROWS = 1_048_575
NORMALIZED_SPOOL_MAGIC = b"AISNRM1\0"
NORMALIZED_RECORD = struct.Struct("<ddQIIBq16s")
NORMALIZED_COUNTS = struct.Struct("<32Q")
HISTORICAL_MESSAGE_TYPES = (1, 3, 4, 18, 19)
PERIOD_A = "period_a"
PERIOD_B = "period_b"
LOGICAL_DAY = "logical_day"
REVIEW_WORKBOOK_SCHEMA = "AIS_V15_REVIEW_1"

DIRECTION_ORDER = [
    "北", "北微東", "北北東", "東北微北", "東北", "東北微東", "東北東", "東微北", "東",
    "東微南", "東南東", "東南微南", "東南", "東南微西", "南南東", "南微東", "南",
    "南微西", "南南西", "西南微南", "西南", "西南微西", "西南西", "西微南", "西",
    "西微北", "西北西", "西北微西", "西北", "西北微北", "北北西", "北微西",
]

LEGACY_DIRECTION_ABBREVIATIONS = [
    "N", "NbE", "NNE", "NEbN", "NE", "NEbE", "ENE", "EbN", "E", "EbS", "ESE",
    "SEbE", "SE", "SEbS", "SSE", "SbE", "S", "SbW", "SSW", "SWbS", "SW", "SWbW",
    "WSW", "WbS", "W", "WbN", "WNW", "NWbW", "NW", "NWbN", "NNW", "NbW",
]

# 影片所示：北至東南東，以及西南西至北微西，共 21 個海向方位。
OPEN_SEA_INDEXES = tuple(range(0, 11)) + tuple(range(22, 32))
OPEN_SEA_DIRECTIONS = tuple(DIRECTION_ORDER[index] for index in OPEN_SEA_INDEXES)
COASTAL_REVIEW_DIRECTIONS = {"西南西", "西微南", "西"}

FILENAME_RE = re.compile(
    r"^D&TMOK[ \t]+(?P<port>[A-Z][A-Z0-9]{1,15})_(?P<date>\d{8})(?:_(?P<suffix>[A-Z0-9_-]+))?\.xlsx$",
    re.IGNORECASE,
)
PORT_RE = re.compile(r"^[A-Z][A-Z0-9]{1,15}$")

ProgressCallback = Callable[[dict], None]


class CancelledError(RuntimeError):
    pass


class SourceFileError(RuntimeError):
    pass


@dataclass(frozen=True)
class SourceFileIdentity:
    port: str
    day: date
    suffix: str | None = None


@dataclass(frozen=True)
class SourceFragment:
    port: str
    day: date
    suffix: str | None
    path: Path


@dataclass(frozen=True)
class DayInput:
    port: str
    day: date
    fragments: tuple[SourceFragment, ...]


@dataclass(frozen=True)
class Candidate:
    distance: float
    bearing: float
    source_row: int
    rank: int | None = None
    candidate_id: str = ""
    timestamp: datetime | None = None
    mmsi: int | None = None
    channel: str | None = None
    fragment: str = ""
    sheet: str = ""
    msg_type: int | None = None


@dataclass(frozen=True, slots=True)
class NormalizedRecord:
    distance: float
    bearing: float
    source_key: int
    second_of_day: int
    msg_type: int
    mmsi: int | None
    channel: str

    @property
    def period(self) -> str:
        return PERIOD_A if self.second_of_day < 12 * 60 * 60 else PERIOD_B


@dataclass(frozen=True)
class FragmentInfo:
    path: Path
    suffix: str | None
    sha256: str
    sheet: str
    rows_scanned: int
    min_second: int | None
    max_second: int | None
    occupied_seconds: int
    aliases: tuple[Path, ...] = ()


@dataclass
class DirectionResult:
    direction: str
    candidates: list[Candidate] = field(default_factory=list)
    selected: float | None = None
    selected_bearing: float | None = None
    selected_rank: int | None = None
    cluster_count: int = 0
    status: str = "無資料"
    reason: str = "找不到符合條件的資料"
    over_cap_count: int = 0
    over_cap_max: float | None = None
    selected_candidate_id: str | None = None


@dataclass
class DayResult:
    day: date
    source_fragments: tuple[Path, ...]
    directions: dict[str, DirectionResult]
    period_directions: dict[str, dict[str, DirectionResult]] = field(default_factory=dict)
    fragment_info: tuple[FragmentInfo, ...] = ()
    diagnostics: tuple[str, ...] = ()
    rows_scanned: int = 0
    rows_accepted: int = 0
    rows_invalid: int = 0
    rows_wrong_message: int = 0
    rows_not_east: int = 0
    rows_legacy: int = 0
    rows_normalized: int = 0
    spool_sha256: str = ""
    elapsed_seconds: float = 0.0
    note: str = ""

    @property
    def source_file(self) -> Path | None:
        """Deprecated single-file view; multi-fragment days deliberately return None."""
        return self.source_fragments[0] if len(self.source_fragments) == 1 else None


@dataclass(frozen=True)
class ReviewDecision:
    scope: str
    day: date
    direction: str
    candidate_id: str | None
    forced_numeric: float | None = None
    note: str = ""


@dataclass(frozen=True)
class DecisionSnapshot:
    port: str
    year: int
    month: int
    source_manifest_hash: str
    decisions: tuple[ReviewDecision, ...]

    def by_key(self) -> dict[tuple[str, date, str], ReviewDecision]:
        return {
            (decision.scope, decision.day, decision.direction): decision
            for decision in self.decisions
        }


@dataclass
class AppConfig:
    input_dir: Path
    output_path: Path
    port: str
    year: int
    month: int
    legacy_output_path: Path | None = None
    max_distance: float = 500.0
    tolerance: float = 0.10
    cluster_size: int = 3
    top_candidates: int = 50
    message_types: tuple[int, ...] = (1, 2, 3, 18, 19)
    workers: int = 1
    overwrite: bool = False
    max_files: int | None = None

    def __post_init__(self) -> None:
        self.port = normalize_port(self.port)

    def validate(self) -> None:
        if not self.input_dir.is_dir():
            raise ValueError(f"來源資料夾不存在：{self.input_dir}")
        if not 2000 <= self.year <= 2100:
            raise ValueError("年份必須介於 2000 到 2100。")
        if not 1 <= self.month <= 12:
            raise ValueError("月份必須介於 1 到 12。")
        if not 0 < self.max_distance <= 5000:
            raise ValueError("距離上限必須大於 0。")
        if not 0 < self.tolerance < 1:
            raise ValueError("群聚容許差必須介於 0% 到 100% 之間。")
        if not 2 <= self.cluster_size <= 20:
            raise ValueError("群聚筆數必須介於 2 到 20。")
        if not self.cluster_size <= self.top_candidates <= 1000:
            raise ValueError("候選保留筆數不可小於群聚筆數。")
        if not self.message_types:
            raise ValueError("至少要指定一個 AIS 訊息類型。")
        if not 1 <= self.workers <= 8:
            raise ValueError("平行檔案數必須介於 1 到 8。")
        if self.output_path.suffix.lower() != ".xlsx":
            raise ValueError("輸出檔必須是 .xlsx。")
        if self.legacy_output_path is not None and self.legacy_output_path.suffix.lower() != ".xlsx":
            raise ValueError("原格式相容版輸出檔必須是 .xlsx。")
        if self.legacy_output_path is not None and self.legacy_output_path.resolve() == self.output_path.resolve():
            raise ValueError("新版與原格式版不可使用同一個輸出檔名。")
        if self.output_path.exists() and not self.overwrite:
            raise FileExistsError(f"輸出檔已存在：{self.output_path}")
        if self.legacy_output_path is not None and self.legacy_output_path.exists() and not self.overwrite:
            raise FileExistsError(f"原格式相容版已存在：{self.legacy_output_path}")


@dataclass(frozen=True)
class ProcessingJob:
    port: str
    year: int
    month: int
    days: tuple[DayInput, ...]
    warnings: tuple[str, ...] = ()

    @property
    def files(self) -> tuple[tuple[date, Path], ...]:
        """Flattened compatibility view. Internal processing must use ``days``."""
        return tuple(
            (day_input.day, fragment.path)
            for day_input in self.days
            for fragment in day_input.fragments
        )


def emit(callback: ProgressCallback | None, **payload: object) -> None:
    if callback:
        callback(payload)


def default_worker_count() -> int:
    return min(3, max(1, os.cpu_count() or 1))


def default_output_directory() -> Path:
    if getattr(sys, "frozen", False):
        application_root = Path(sys.executable).resolve().parent
    else:
        application_root = Path(__file__).resolve().parent.parent
    return application_root / "output" / "AIS月報"


def normalize_port(port: str) -> str:
    normalized = port.strip().upper()
    if not PORT_RE.fullmatch(normalized):
        raise ValueError("港別代碼必須是 2–16 位英數字，且第一個字元必須是英文字母。")
    return normalized


def parse_source_filename(filename: str) -> SourceFileIdentity | None:
    match = FILENAME_RE.match(filename)
    if not match:
        return None
    try:
        parsed_day = datetime.strptime(match.group("date"), "%Y%m%d").date()
    except ValueError:
        return None
    suffix = match.group("suffix")
    return SourceFileIdentity(
        port=match.group("port").upper(),
        day=parsed_day,
        suffix=suffix.upper() if suffix else None,
    )


def parse_source_date(filename: str) -> date | None:
    parsed = parse_source_filename(filename)
    return parsed.day if parsed is not None else None


def scan_source_files(folder: Path) -> tuple[dict[str, dict[date, DayInput]], dict[str, list[str]]]:
    grouped: dict[str, dict[date, list[SourceFragment]]] = {}
    warnings: dict[str, list[str]] = {}
    for path in sorted(folder.glob("*.xlsx"), key=lambda item: item.name.casefold()):
        identity = parse_source_filename(path.name)
        if identity is None:
            continue
        grouped.setdefault(identity.port, {}).setdefault(identity.day, []).append(
            SourceFragment(
                port=identity.port,
                day=identity.day,
                suffix=identity.suffix,
                path=path,
            )
        )

    catalog: dict[str, dict[date, DayInput]] = {}
    for port, port_days in grouped.items():
        output_days: dict[date, DayInput] = {}
        port_warnings = warnings.setdefault(port, [])
        for parsed_day, fragments in port_days.items():
            ordered = tuple(sorted(fragments, key=lambda item: item.path.name.casefold()))
            output_days[parsed_day] = DayInput(port=port, day=parsed_day, fragments=ordered)
            if len(ordered) > 1:
                port_warnings.append(
                    f"{port} {parsed_day:%Y-%m-%d} 將合併 {len(ordered)} 個來源分片："
                    + "、".join(fragment.path.name for fragment in ordered)
                )
        catalog[port] = output_days
    return catalog, warnings


def discover_source_files(folder: Path, port: str | None = None) -> tuple[dict[date, DayInput], list[str]]:
    catalog, warnings = scan_source_files(folder)
    if port is None:
        if len(catalog) > 1:
            listing = "、".join(sorted(catalog))
            raise ValueError(f"偵測到多個港別：{listing}；請先指定港別。")
        if not catalog:
            return {}, []
        normalized_port = next(iter(catalog))
    else:
        normalized_port = normalize_port(port)
    return catalog.get(normalized_port, {}), warnings.get(normalized_port, [])


def detect_source_catalog(folder: Path) -> dict[str, list[tuple[int, int, int]]]:
    catalog, _ = scan_source_files(folder)
    result: dict[str, list[tuple[int, int, int]]] = {}
    for port, files in sorted(catalog.items()):
        counts: dict[tuple[int, int], int] = {}
        for parsed in files:
            key = (parsed.year, parsed.month)
            counts[key] = counts.get(key, 0) + 1
        result[port] = [
            (year, month, count)
            for (year, month), count in sorted(counts.items())
        ]
    return result


def detect_available_months(folder: Path, port: str | None = None) -> list[tuple[int, int, int]]:
    catalog = detect_source_catalog(folder)
    if port is None:
        if len(catalog) > 1:
            listing = "、".join(catalog)
            raise ValueError(f"偵測到多個港別：{listing}；請先指定港別。")
        return next(iter(catalog.values()), [])
    return catalog.get(normalize_port(port), [])


def month_option_label(year: int, month: int, count: int) -> str:
    return f"{year} 年 {month} 月（{count} 個每日檔）"


def resolve_source_period(
    folder: Path,
    port: str | None = None,
    year: int | None = None,
    month: int | None = None,
) -> tuple[str, int, int, int]:
    """Resolve one unambiguous port/month processing job from source filenames."""
    if (year is None) != (month is None):
        raise ValueError("年份與月份必須同時指定；也可以兩者都省略，改由檔名自動判定。")
    catalog = detect_source_catalog(folder)
    if not catalog:
        raise ValueError("找不到符合 D&TMOK <PORT>_YYYYMMDD_*.xlsx 格式的檔案。")
    if port is None:
        if len(catalog) != 1:
            listing = "、".join(catalog)
            raise ValueError(f"來源資料夾包含多個港別：{listing}；請以 --port 指定其中一個港別。")
        selected_port = next(iter(catalog))
    else:
        selected_port = normalize_port(port)
        if selected_port not in catalog:
            listing = "、".join(catalog)
            raise ValueError(f"檔名中沒有港別 {selected_port}；偵測到：{listing}。")

    available = catalog[selected_port]
    if year is None:
        selected_year, selected_month, count = max(available, key=lambda item: (item[0], item[1]))
        return selected_port, selected_year, selected_month, count
    for detected in available:
        if detected[:2] == (year, month):
            return selected_port, *detected
    listing = "、".join(f"{item[0]} 年 {item[1]} 月" for item in available)
    raise ValueError(f"港別 {selected_port} 的檔名中沒有 {year} 年 {month} 月資料；偵測到：{listing}。")


def resolve_source_month(
    folder: Path,
    year: int | None = None,
    month: int | None = None,
    port: str | None = None,
) -> tuple[int, int, int]:
    """Resolve the processing month from fixed-format source filenames."""
    _port, selected_year, selected_month, count = resolve_source_period(folder, port, year, month)
    return selected_year, selected_month, count


def degree_to_direction_index(value: object) -> tuple[int, float] | None:
    try:
        degree = float(value)
    except (TypeError, ValueError):
        return None
    if not math.isfinite(degree):
        return None
    normalized = degree % 360.0
    index = min(int(normalized // 11.25), 31)
    return index, normalized


def _number(value: object) -> float | None:
    try:
        result = float(value)
    except (TypeError, ValueError):
        return None
    return result if math.isfinite(result) else None


def _message_type(value: object) -> int | None:
    number = _number(value)
    if number is None or not number.is_integer():
        return None
    return int(number)


def _normalized_header(value: object) -> str:
    text = "" if value is None else str(value)
    return re.sub(r"\s+", " ", text.strip().casefold())


REQUIRED_HEADERS = {
    "year": {"year"},
    "month": {"month"},
    "day": {"day"},
    "hour": {"hour"},
    "minute": {"minute"},
    "second": {"second"},
    "channel": {"channel"},
    "msg_type": {"msg_type", "msg type"},
    "mmsi": {"mmsi"},
    "longitude_desc": {"longitude_desc", "longitude desc"},
    "bearing": {"bearing"},
    "distance": {"distance in nautical miles", "distance_in_nautical_miles"},
}


def locate_headers(header_row: Sequence[object]) -> dict[str, int]:
    normalized = [_normalized_header(value) for value in header_row]
    result: dict[str, int] = {}
    for canonical, aliases in REQUIRED_HEADERS.items():
        for index, value in enumerate(normalized):
            if value in aliases:
                result[canonical] = index
                break
    missing = [key for key in REQUIRED_HEADERS if key not in result]
    if missing:
        friendly = ", ".join(missing)
        raise ValueError(f"缺少必要欄位：{friendly}")
    return result


def locate_data_worksheet(workbook: object) -> tuple[object, dict[str, int]]:
    worksheets = list(workbook.worksheets)
    worksheets.sort(key=lambda worksheet: worksheet.title.strip().casefold() != "ais")
    inspected: list[str] = []
    for worksheet in worksheets:
        header = next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if not header:
            inspected.append(f"{worksheet.title}（空白）")
            continue
        try:
            return worksheet, locate_headers(header)
        except ValueError:
            inspected.append(worksheet.title)
    names = "、".join(inspected) if inspected else "無工作表"
    raise ValueError(f"找不到包含必要欄位的工作表（已檢查：{names}）")


def _required_integer(value: object, label: str, source_file: Path, excel_row: int) -> int:
    number = _number(value)
    if number is None or not number.is_integer():
        raise SourceFileError(
            f"來源檔「{source_file.name}」第 {excel_row:,} 列的 {label} 不是有效整數。"
        )
    return int(number)


def _file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while chunk := handle.read(4 * 1024 * 1024):
            digest.update(chunk)
    return digest.hexdigest()


def _candidate_id(fragment: FragmentInfo, record: NormalizedRecord) -> str:
    source_row = record.source_key & 0xFFFFFFFF
    payload = "|".join(
        (
            fragment.sha256,
            fragment.sheet,
            str(source_row),
            str(record.second_of_day),
            str(record.msg_type),
            "" if record.mmsi is None else str(record.mmsi),
            record.channel,
            format(record.bearing, ".17g"),
            format(record.distance, ".17g"),
        )
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:24]


def select_cluster(
    direction: str,
    candidates: list[Candidate],
    tolerance: float,
    cluster_size: int,
    over_cap_count: int,
    over_cap_max: float | None,
) -> DirectionResult:
    result = DirectionResult(
        direction=direction,
        candidates=candidates,
        over_cap_count=over_cap_count,
        over_cap_max=over_cap_max,
    )
    if not candidates:
        if over_cap_count:
            result.reason = f"只有超過距離上限的資料（{over_cap_count:,} 筆）"
        return result

    for start, candidate in enumerate(candidates):
        floor = candidate.distance * (1.0 - tolerance)
        count = 0
        for follower in candidates[start:]:
            if follower.distance + 1e-12 < floor:
                break
            count += 1
        if count >= cluster_size:
            result.selected = candidate.distance
            result.selected_bearing = candidate.bearing
            result.selected_rank = start + 1
            result.cluster_count = count
            result.status = "自動採用"
            discarded = start
            result.reason = (
                f"第 {start + 1} 名起有 {count} 筆落在 {tolerance:.0%} 範圍內"
                + (f"；前 {discarded} 筆視為孤立值" if discarded else "")
            )
            if direction in COASTAL_REVIEW_DIRECTIONS and candidate.distance > 10:
                result.status = "待複核"
                result.reason += "；此岸向方位超過影片提示的 10 NM 經驗值"
            return result

    fallback = candidates[0]
    result.selected = fallback.distance
    result.selected_bearing = fallback.bearing
    result.selected_rank = 1
    result.cluster_count = 1
    result.status = "待複核"
    result.reason = f"前 {len(candidates)} 筆中找不到至少 {cluster_size} 筆的 {tolerance:.0%} 群聚，暫用最高值"
    return result


def select_cluster_from_sorted_rows(
    direction: str,
    rows: list[tuple[float, int, float]],
    config: AppConfig,
    over_cap_count: int,
    over_cap_max: float | None,
) -> DirectionResult:
    valid_start = over_cap_count
    candidates = [
        Candidate(distance=item[0], source_row=item[1], bearing=item[2], rank=rank)
        for rank, item in enumerate(
            rows[valid_start : valid_start + config.top_candidates],
            start=1,
        )
    ]
    result = DirectionResult(
        direction=direction,
        candidates=candidates,
        over_cap_count=over_cap_count,
        over_cap_max=over_cap_max,
    )
    valid_count = len(rows) - valid_start
    if valid_count <= 0:
        if over_cap_count:
            result.reason = f"只有超過距離上限的資料（{over_cap_count:,} 筆）"
        return result

    last_possible = len(rows) - config.cluster_size
    for start in range(valid_start, last_possible + 1):
        candidate_distance, _source_row, candidate_bearing = rows[start]
        floor = candidate_distance * (1.0 - config.tolerance)
        required_follower = rows[start + config.cluster_size - 1]
        if required_follower[0] + 1e-12 < floor:
            continue
        stop = start + config.cluster_size
        while stop < len(rows) and rows[stop][0] + 1e-12 >= floor:
            stop += 1
        rank = start - valid_start + 1
        result.selected = candidate_distance
        result.selected_bearing = candidate_bearing
        result.selected_rank = rank
        result.cluster_count = stop - start
        result.status = "自動採用"
        result.reason = (
            f"第 {rank} 名起有 {result.cluster_count} 筆落在 {config.tolerance:.0%} 範圍內"
            + (f"；前 {rank - 1} 筆視為孤立值" if rank > 1 else "")
        )
        if direction in COASTAL_REVIEW_DIRECTIONS and candidate_distance > 10:
            result.status = "待複核"
            result.reason += "；此岸向方位超過影片提示的 10 NM 經驗值"
        if rank > config.top_candidates:
            extra_stop = min(stop, start + config.top_candidates)
            candidates.extend(
                Candidate(
                    distance=item[0],
                    source_row=item[1],
                    bearing=item[2],
                    rank=actual_rank,
                )
                for actual_rank, item in enumerate(rows[start:extra_stop], start=rank)
            )
        return result

    fallback_distance, _source_row, fallback_bearing = rows[valid_start]
    result.selected = fallback_distance
    result.selected_bearing = fallback_bearing
    result.selected_rank = 1
    result.cluster_count = 1
    result.status = "待複核"
    result.reason = f"全部 {valid_count:,} 筆中找不到至少 {config.cluster_size} 筆的 {config.tolerance:.0%} 群聚，暫用最高值"
    return result


def select_cluster_from_normalized_records(
    direction: str,
    records: list[NormalizedRecord],
    config: AppConfig,
    fragment_info: Sequence[FragmentInfo],
    parsed_day: date,
) -> DirectionResult:
    over_cap_count = 0
    over_cap_max: float | None = None
    for record in records:
        if record.distance <= config.max_distance:
            break
        over_cap_count += 1
        if over_cap_max is None:
            over_cap_max = record.distance
    raw_result = select_cluster_from_sorted_rows(
        direction,
        [(record.distance, record.source_key, record.bearing) for record in records],
        config,
        over_cap_count,
        over_cap_max,
    )
    wanted_keys = {candidate.source_row for candidate in raw_result.candidates}
    records_by_key = {
        record.source_key: record for record in records if record.source_key in wanted_keys
    }
    enriched: list[Candidate] = []
    for candidate in raw_result.candidates:
        record = records_by_key[candidate.source_row]
        fragment_index = record.source_key >> 32
        source_row = record.source_key & 0xFFFFFFFF
        fragment = fragment_info[fragment_index]
        enriched.append(
            Candidate(
                distance=record.distance,
                bearing=record.bearing,
                source_row=source_row,
                rank=candidate.rank,
                candidate_id=_candidate_id(fragment, record),
                timestamp=datetime.combine(parsed_day, datetime.min.time()).replace(
                    hour=record.second_of_day // 3600,
                    minute=(record.second_of_day % 3600) // 60,
                    second=record.second_of_day % 60,
                ),
                mmsi=record.mmsi,
                channel=record.channel,
                fragment=fragment.path.name,
                sheet=fragment.sheet,
                msg_type=record.msg_type,
            )
        )
    raw_result.candidates = enriched
    if raw_result.selected_rank is not None:
        selected_index = over_cap_count + raw_result.selected_rank - 1
        if 0 <= selected_index < len(records):
            selected_record = records[selected_index]
            selected_fragment = fragment_info[selected_record.source_key >> 32]
            raw_result.selected_candidate_id = _candidate_id(selected_fragment, selected_record)
    return raw_result


def empty_day_result(day: date, note: str = "當日來源檔缺漏") -> DayResult:
    directions = {
        name: DirectionResult(direction=name, status="無資料", reason=note)
        for name in DIRECTION_ORDER
    }
    period_directions = {
        period: {
            name: DirectionResult(direction=name, status="無資料", reason=note)
            for name in DIRECTION_ORDER
        }
        for period in (PERIOD_A, PERIOD_B)
    }
    return DayResult(
        day=day,
        source_fragments=(),
        directions=directions,
        period_directions=period_directions,
        note=note,
    )


def _job_storage_token(config: AppConfig) -> str:
    return f"{config.port}_{config.year}_{config.month:02d}"


def _cache_directory(config: AppConfig) -> Path:
    return config.output_path.parent / f".{config.output_path.stem}_{_job_storage_token(config)}_cache"


def _normalized_spool_directory(config: AppConfig) -> Path:
    return config.output_path.parent / f".{config.output_path.stem}_{_job_storage_token(config)}_rows"


def _legacy_spool_directory(config: AppConfig) -> Path:
    """Compatibility alias for the v1.5 normalized logical-day spool."""
    return _normalized_spool_directory(config)


def _normalized_spool_path(config: AppConfig, parsed_day: date) -> Path:
    return _normalized_spool_directory(config) / f"{parsed_day:%Y%m%d}.bin"


def _legacy_spool_path(config: AppConfig, parsed_day: date) -> Path:
    return _normalized_spool_path(config, parsed_day)


def _cache_signature(day_input: DayInput, config: AppConfig) -> dict[str, object]:
    return {
        "cache_version": CACHE_VERSION,
        "sources": [
            {
                "path": str(fragment.path.resolve()),
                "size": fragment.path.stat().st_size,
                "modified_ns": fragment.path.stat().st_mtime_ns,
                "suffix": fragment.suffix,
            }
            for fragment in day_input.fragments
        ],
        "port": config.port,
        "year": config.year,
        "month": config.month,
        "date": day_input.day.isoformat(),
        "max_distance": config.max_distance,
        "tolerance": config.tolerance,
        "cluster_size": config.cluster_size,
        "top_candidates": config.top_candidates,
        "message_types": list(config.message_types),
        "normalized_spool_version": NORMALIZED_SPOOL_VERSION,
    }


def _direction_result_to_dict(result: DirectionResult) -> dict[str, object]:
    return {
        "selected": result.selected,
        "selected_bearing": result.selected_bearing,
        "selected_rank": result.selected_rank,
        "selected_candidate_id": result.selected_candidate_id,
        "cluster_count": result.cluster_count,
        "status": result.status,
        "reason": result.reason,
        "over_cap_count": result.over_cap_count,
        "over_cap_max": result.over_cap_max,
        "candidates": [
            {
                "distance": item.distance,
                "bearing": item.bearing,
                "source_row": item.source_row,
                "rank": item.rank,
                "candidate_id": item.candidate_id,
                "timestamp": item.timestamp.isoformat() if item.timestamp else None,
                "mmsi": item.mmsi,
                "channel": item.channel,
                "fragment": item.fragment,
                "sheet": item.sheet,
                "msg_type": item.msg_type,
            }
            for item in result.candidates
        ],
    }


def _direction_result_from_dict(direction: str, raw: object) -> DirectionResult:
    if not isinstance(raw, dict):
        raise ValueError("快取 direction 格式錯誤")
    raw_candidates = raw.get("candidates", [])
    candidates = [
        Candidate(
            distance=float(item["distance"]),
            bearing=float(item["bearing"]),
            source_row=int(item["source_row"]),
            rank=int(item["rank"]) if item.get("rank") is not None else None,
            candidate_id=str(item.get("candidate_id", "")),
            timestamp=(
                datetime.fromisoformat(str(item["timestamp"]))
                if item.get("timestamp")
                else None
            ),
            mmsi=int(item["mmsi"]) if item.get("mmsi") is not None else None,
            channel=str(item["channel"]) if item.get("channel") is not None else None,
            fragment=str(item.get("fragment", "")),
            sheet=str(item.get("sheet", "")),
            msg_type=int(item["msg_type"]) if item.get("msg_type") is not None else None,
        )
        for item in raw_candidates
    ]
    return DirectionResult(
        direction=direction,
        candidates=candidates,
        selected=raw.get("selected"),
        selected_bearing=raw.get("selected_bearing"),
        selected_rank=raw.get("selected_rank"),
        selected_candidate_id=raw.get("selected_candidate_id"),
        cluster_count=int(raw.get("cluster_count", 0)),
        status=str(raw.get("status", "無資料")),
        reason=str(raw.get("reason", "")),
        over_cap_count=int(raw.get("over_cap_count", 0)),
        over_cap_max=raw.get("over_cap_max"),
    )


def _day_result_to_dict(day_result: DayResult) -> dict[str, object]:
    return {
        "day": day_result.day.isoformat(),
        "source_fragments": [str(path) for path in day_result.source_fragments],
        "fragment_info": [
            {
                "path": str(info.path),
                "suffix": info.suffix,
                "sha256": info.sha256,
                "sheet": info.sheet,
                "rows_scanned": info.rows_scanned,
                "min_second": info.min_second,
                "max_second": info.max_second,
                "occupied_seconds": info.occupied_seconds,
                "aliases": [str(path) for path in info.aliases],
            }
            for info in day_result.fragment_info
        ],
        "diagnostics": list(day_result.diagnostics),
        "rows_scanned": day_result.rows_scanned,
        "rows_accepted": day_result.rows_accepted,
        "rows_invalid": day_result.rows_invalid,
        "rows_wrong_message": day_result.rows_wrong_message,
        "rows_not_east": day_result.rows_not_east,
        "rows_legacy": day_result.rows_legacy,
        "rows_normalized": day_result.rows_normalized,
        "spool_sha256": day_result.spool_sha256,
        "elapsed_seconds": day_result.elapsed_seconds,
        "note": day_result.note,
        "directions": {
            direction: _direction_result_to_dict(result)
            for direction, result in day_result.directions.items()
        },
        "period_directions": {
            period: {
                direction: _direction_result_to_dict(result)
                for direction, result in directions.items()
            }
            for period, directions in day_result.period_directions.items()
        },
    }


def _day_result_from_dict(payload: dict[str, object]) -> DayResult:
    directions: dict[str, DirectionResult] = {}
    raw_directions = payload["directions"]
    if not isinstance(raw_directions, dict):
        raise ValueError("快取 directions 格式錯誤")
    for direction, raw in raw_directions.items():
        directions[str(direction)] = _direction_result_from_dict(str(direction), raw)
    period_directions: dict[str, dict[str, DirectionResult]] = {}
    raw_periods = payload.get("period_directions", {})
    if not isinstance(raw_periods, dict):
        raise ValueError("快取 period_directions 格式錯誤")
    for period, raw_period_directions in raw_periods.items():
        if not isinstance(raw_period_directions, dict):
            raise ValueError("快取 period direction 格式錯誤")
        period_directions[str(period)] = {
            str(direction): _direction_result_from_dict(str(direction), raw)
            for direction, raw in raw_period_directions.items()
        }
    source_values = payload.get("source_fragments")
    if source_values is None:
        source_value = payload.get("source_file")
        source_values = [source_value] if source_value else []
    if not isinstance(source_values, list):
        raise ValueError("快取 source_fragments 格式錯誤")
    return DayResult(
        day=date.fromisoformat(str(payload["day"])),
        source_fragments=tuple(Path(str(value)) for value in source_values),
        directions=directions,
        period_directions=period_directions,
        fragment_info=tuple(
            FragmentInfo(
                path=Path(str(item["path"])),
                suffix=str(item["suffix"]) if item.get("suffix") is not None else None,
                sha256=str(item["sha256"]),
                sheet=str(item["sheet"]),
                rows_scanned=int(item["rows_scanned"]),
                min_second=int(item["min_second"]) if item.get("min_second") is not None else None,
                max_second=int(item["max_second"]) if item.get("max_second") is not None else None,
                occupied_seconds=int(item["occupied_seconds"]),
                aliases=tuple(Path(str(value)) for value in item.get("aliases", [])),
            )
            for item in payload.get("fragment_info", [])
        ),
        diagnostics=tuple(str(value) for value in payload.get("diagnostics", [])),
        rows_scanned=int(payload.get("rows_scanned", 0)),
        rows_accepted=int(payload.get("rows_accepted", 0)),
        rows_invalid=int(payload.get("rows_invalid", 0)),
        rows_wrong_message=int(payload.get("rows_wrong_message", 0)),
        rows_not_east=int(payload.get("rows_not_east", 0)),
        rows_legacy=int(payload.get("rows_legacy", payload.get("rows_accepted", 0))),
        rows_normalized=int(payload.get("rows_normalized", payload.get("rows_legacy", 0))),
        spool_sha256=str(payload.get("spool_sha256", "")),
        elapsed_seconds=float(payload.get("elapsed_seconds", 0.0)),
        note=str(payload.get("note", "")),
    )


def load_cached_day(day_input: DayInput, config: AppConfig) -> DayResult | None:
    cache_path = _cache_directory(config) / f"{day_input.day:%Y%m%d}.json"
    if not cache_path.is_file():
        return None
    try:
        payload = json.loads(cache_path.read_text(encoding="utf-8"))
        if payload.get("signature") != _cache_signature(day_input, config):
            return None
        result = _day_result_from_dict(payload["result"])
        spool_path = _normalized_spool_path(config, day_input.day)
        if not spool_path.is_file() or sum(normalized_spool_counts(spool_path)) != result.rows_normalized:
            return None
        if result.spool_sha256 and _file_sha256(spool_path) != result.spool_sha256:
            return None
        return result
    except (OSError, ValueError, TypeError, KeyError, json.JSONDecodeError):
        return None


def save_cached_day(day_input: DayInput, config: AppConfig, day_result: DayResult) -> None:
    cache_dir = _cache_directory(config)
    cache_dir.mkdir(parents=True, exist_ok=True)
    cache_path = cache_dir / f"{day_input.day:%Y%m%d}.json"
    temp_path = cache_path.with_suffix(".tmp")
    payload = {
        "signature": _cache_signature(day_input, config),
        "result": _day_result_to_dict(day_result),
    }
    temp_path.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    os.replace(temp_path, cache_path)


def write_normalized_spool(
    config: AppConfig,
    parsed_day: date,
    rows: list[list[NormalizedRecord]],
) -> tuple[Path, str]:
    spool_dir = _normalized_spool_directory(config)
    spool_dir.mkdir(parents=True, exist_ok=True)
    target = _normalized_spool_path(config, parsed_day)
    temporary = target.with_suffix(".tmp")
    counts = [len(bucket) for bucket in rows]
    with temporary.open("wb") as handle:
        handle.write(NORMALIZED_SPOOL_MAGIC)
        handle.write(NORMALIZED_COUNTS.pack(*counts))
        for bucket in rows:
            for record in bucket:
                encoded_channel = record.channel.encode("utf-8")
                if len(encoded_channel) > 16:
                    raise SourceFileError(
                        f"channel 值過長，無法寫入 normalized spool：{record.channel!r}"
                    )
                handle.write(
                    NORMALIZED_RECORD.pack(
                        record.distance,
                        record.bearing,
                        record.source_key,
                        record.second_of_day,
                        record.msg_type,
                        0 if record.period == PERIOD_A else 1,
                        -1 if record.mmsi is None else record.mmsi,
                        encoded_channel.ljust(16, b"\0"),
                    )
                )
        handle.flush()
        os.fsync(handle.fileno())
    digest = _file_sha256(temporary)
    os.replace(temporary, target)
    return target, digest


def normalized_spool_counts(path: Path) -> list[int]:
    with path.open("rb") as handle:
        if handle.read(len(NORMALIZED_SPOOL_MAGIC)) != NORMALIZED_SPOOL_MAGIC:
            raise ValueError(f"normalized spool 格式錯誤：{path.name}")
        counts_raw = handle.read(NORMALIZED_COUNTS.size)
        if len(counts_raw) != NORMALIZED_COUNTS.size:
            raise ValueError(f"normalized spool 不完整：{path.name}")
        counts = list(NORMALIZED_COUNTS.unpack(counts_raw))
    expected_size = (
        len(NORMALIZED_SPOOL_MAGIC)
        + NORMALIZED_COUNTS.size
        + sum(counts) * NORMALIZED_RECORD.size
    )
    if path.stat().st_size != expected_size:
        raise ValueError(f"normalized spool 大小不符：{path.name}")
    return counts


def read_normalized_spool(
    path: Path,
) -> tuple[list[int], Iterable[tuple[int, NormalizedRecord]]]:
    handle = path.open("rb")
    try:
        counts = normalized_spool_counts(path)
        handle.seek(len(NORMALIZED_SPOOL_MAGIC) + NORMALIZED_COUNTS.size)

        def records() -> Iterable[tuple[int, NormalizedRecord]]:
            try:
                for direction_index, count in enumerate(counts):
                    for _ in range(count):
                        raw = handle.read(NORMALIZED_RECORD.size)
                        if len(raw) != NORMALIZED_RECORD.size:
                            raise ValueError(f"normalized spool 資料中斷：{path.name}")
                        (
                            distance,
                            bearing,
                            source_key,
                            second_of_day,
                            msg_type,
                            _period_code,
                            mmsi,
                            raw_channel,
                        ) = NORMALIZED_RECORD.unpack(raw)
                        yield direction_index, NormalizedRecord(
                            distance=distance,
                            bearing=bearing,
                            source_key=source_key,
                            second_of_day=second_of_day,
                            msg_type=msg_type,
                            mmsi=None if mmsi < 0 else mmsi,
                            channel=raw_channel.rstrip(b"\0").decode("utf-8"),
                        )
                if handle.read(1):
                    raise ValueError(f"normalized spool 尾端有非預期資料：{path.name}")
            finally:
                handle.close()

        return counts, records()
    except Exception:
        handle.close()
        raise


def legacy_spool_counts(path: Path) -> list[int]:
    return normalized_spool_counts(path)


def read_legacy_spool(path: Path) -> tuple[list[int], Iterable[tuple[int, float, float]]]:
    counts, normalized_records = read_normalized_spool(path)

    def records() -> Iterable[tuple[int, float, float]]:
        for direction_index, record in normalized_records:
            yield direction_index, record.distance, record.bearing

    return counts, records()


def _process_day_worker(arguments: tuple[DayInput, AppConfig]) -> tuple[date, DayResult]:
    day_input, config = arguments
    return day_input.day, process_day_input(day_input, config)


def preflight_output_space(config: AppConfig, source_files: Iterable[Path]) -> tuple[int, int | None]:
    targets = [config.output_path]
    if config.legacy_output_path is not None:
        targets.append(config.legacy_output_path)
    for target in targets:
        target.parent.mkdir(parents=True, exist_ok=True)

    source_size = sum(path.stat().st_size for path in source_files)
    if config.legacy_output_path is None:
        required = max(256 * 1024**2, int(source_size * 0.08))
    else:
        required = max(1536 * 1024**2, int(source_size * 0.55))
    try:
        free = shutil.disk_usage(config.output_path.parent).free
    except OSError:
        free = None
    if free is not None and free > 0 and free < required:
        raise OSError(
            f"輸出磁碟空間可能不足。估計至少需要 {required / 1024**3:.1f} GB，"
            f"目前可用 {free / 1024**3:.1f} GB。請更換輸出位置或清出空間。"
        )
    return required, free


def process_day_input(
    day_input: DayInput,
    config: AppConfig,
    callback: ProgressCallback | None = None,
    cancel_event: threading.Event | None = None,
) -> DayResult:
    parsed_day = day_input.day
    started = time.perf_counter()
    all_rows: list[list[NormalizedRecord]] = [[] for _ in DIRECTION_ORDER]
    result = DayResult(
        day=parsed_day,
        source_fragments=tuple(fragment.path for fragment in day_input.fragments),
        directions={},
    )

    hashed_fragments = [(fragment, _file_sha256(fragment.path)) for fragment in day_input.fragments]
    by_digest: dict[str, list[SourceFragment]] = {}
    for fragment, digest in hashed_fragments:
        by_digest.setdefault(digest, []).append(fragment)
    canonical_fragments = [
        (digest, sorted(fragments, key=lambda item: item.path.name.casefold()))
        for digest, fragments in by_digest.items()
    ]
    canonical_fragments.sort(key=lambda item: item[0])
    fragment_info: list[FragmentInfo] = []
    occupied_by_fragment: list[bytearray] = []
    diagnostics: list[str] = []

    for fragment_index, (fragment_digest, equivalent_fragments) in enumerate(canonical_fragments):
        fragment = equivalent_fragments[0]
        aliases = tuple(item.path for item in equivalent_fragments[1:])
        if aliases:
            diagnostics.append(
                f"byte-identical duplicate file：{fragment.path.name} 代表 "
                + "、".join(path.name for path in aliases)
            )
        source_file = fragment.path
        try:
            workbook = openpyxl.load_workbook(source_file, read_only=True, data_only=True)
        except (OSError, ValueError, KeyError, zipfile.BadZipFile) as error:
            raise SourceFileError(
                f"無法開啟來源檔「{source_file.name}」。檔案可能損壞、尚未同步完成，或不是有效的 .xlsx。原始錯誤：{error}"
            ) from error
        try:
            try:
                worksheet, indexes = locate_data_worksheet(workbook)
            except ValueError as error:
                raise SourceFileError(
                    f"來源檔「{source_file.name}」{error}。正式五檔輸出需要日期、時間、channel、"
                    "msg_type、mmsi、LONGITUDE_DESC、bearing、distance in nautical miles。"
                ) from error
            max_index = max(indexes.values())
            estimated_rows = max((worksheet.max_row or 1) - 1, 1)
            fragment_rows = 0
            min_second: int | None = None
            max_second: int | None = None
            occupied = bytearray(24 * 60 * 60)

            for excel_row, row in enumerate(worksheet.iter_rows(min_row=2, values_only=True), start=2):
                result.rows_scanned += 1
                fragment_rows += 1
                if cancel_event and excel_row % 10000 == 0 and cancel_event.is_set():
                    raise CancelledError("使用者已取消")
                if len(row) <= max_index:
                    raise SourceFileError(
                        f"來源檔「{source_file.name}」第 {excel_row:,} 列欄位不足，無法建立正式交付 provenance。"
                    )

                year = _required_integer(row[indexes["year"]], "Year", source_file, excel_row)
                month = _required_integer(row[indexes["month"]], "Month", source_file, excel_row)
                day_number = _required_integer(row[indexes["day"]], "Day", source_file, excel_row)
                hour = _required_integer(row[indexes["hour"]], "Hour", source_file, excel_row)
                minute = _required_integer(row[indexes["minute"]], "Minute", source_file, excel_row)
                second = _required_integer(row[indexes["second"]], "Second", source_file, excel_row)
                try:
                    row_day = date(year, month, day_number)
                    datetime(year, month, day_number, hour, minute, second)
                except ValueError as error:
                    raise SourceFileError(
                        f"來源檔「{source_file.name}」第 {excel_row:,} 列日期／時間無效：{error}"
                    ) from error
                if row_day != parsed_day:
                    raise SourceFileError(
                        f"來源檔「{source_file.name}」第 {excel_row:,} 列日期為 {row_day}，"
                        f"不符合檔名 logical day {parsed_day}。"
                    )
                second_of_day = hour * 3600 + minute * 60 + second
                occupied[second_of_day] = 1
                min_second = second_of_day if min_second is None else min(min_second, second_of_day)
                max_second = second_of_day if max_second is None else max(max_second, second_of_day)

                msg_type = _message_type(row[indexes["msg_type"]])
                if msg_type is None or not 0 <= msg_type <= 0xFFFFFFFF:
                    result.rows_invalid += 1
                    continue
                if msg_type not in config.message_types:
                    result.rows_wrong_message += 1
                longitude_desc = row[indexes["longitude_desc"]]
                if str(longitude_desc).strip().casefold() != "east":
                    result.rows_not_east += 1
                    continue

                direction_info = degree_to_direction_index(row[indexes["bearing"]])
                distance = _number(row[indexes["distance"]])
                if direction_info is None or distance is None or distance < 0:
                    result.rows_invalid += 1
                    continue

                direction_index, normalized_bearing = direction_info
                source_key = (fragment_index << 32) | excel_row
                raw_mmsi = _number(row[indexes["mmsi"]])
                mmsi = int(raw_mmsi) if raw_mmsi is not None and raw_mmsi.is_integer() else None
                channel = "" if row[indexes["channel"]] is None else str(row[indexes["channel"]]).strip()
                record = NormalizedRecord(
                    distance=distance,
                    bearing=normalized_bearing,
                    source_key=source_key,
                    second_of_day=second_of_day,
                    msg_type=msg_type,
                    mmsi=mmsi,
                    channel=channel,
                )
                all_rows[direction_index].append(record)
                result.rows_normalized += 1
                if msg_type in HISTORICAL_MESSAGE_TYPES:
                    result.rows_legacy += 1
                if msg_type in config.message_types and distance <= config.max_distance:
                    result.rows_accepted += 1

                if excel_row % 100000 == 0:
                    emit(
                        callback,
                        kind="rows",
                        file=source_file.name,
                        rows=result.rows_scanned,
                        estimated_rows=estimated_rows,
                    )
            fragment_info.append(
                FragmentInfo(
                    path=source_file,
                    suffix=fragment.suffix,
                    sha256=fragment_digest,
                    sheet=worksheet.title,
                    rows_scanned=fragment_rows,
                    min_second=min_second,
                    max_second=max_second,
                    occupied_seconds=sum(occupied),
                    aliases=aliases,
                )
            )
            occupied_by_fragment.append(occupied)
        finally:
            workbook.close()

    for first_index, first in enumerate(fragment_info):
        for second_index in range(first_index + 1, len(fragment_info)):
            second = fragment_info[second_index]
            if (
                first.min_second is not None
                and first.max_second is not None
                and second.min_second is not None
                and second.max_second is not None
                and max(first.min_second, second.min_second) <= min(first.max_second, second.max_second)
            ):
                diagnostics.append(
                    f"time-range overlap：{first.path.name} / {second.path.name}"
                )
            occupied_overlap = sum(
                1
                for left, right in zip(
                    occupied_by_fragment[first_index], occupied_by_fragment[second_index]
                )
                if left and right
            )
            if occupied_overlap:
                diagnostics.append(
                    f"occupied-second overlap：{first.path.name} / {second.path.name} "
                    f"共 {occupied_overlap:,} 秒；未做 timestamp dedupe"
                )

    result.fragment_info = tuple(fragment_info)
    result.diagnostics = tuple(diagnostics)
    for index, direction in enumerate(DIRECTION_ORDER):
        all_rows[index].sort(key=lambda item: (-item.distance, item.source_key))
        logical_records = [
            record for record in all_rows[index] if record.msg_type in config.message_types
        ]
        result.directions[direction] = select_cluster_from_normalized_records(
            direction, logical_records, config, result.fragment_info, parsed_day
        )
    result.period_directions = {}
    for period in (PERIOD_A, PERIOD_B):
        period_results: dict[str, DirectionResult] = {}
        for index, direction in enumerate(DIRECTION_ORDER):
            period_records = [
                record
                for record in all_rows[index]
                if record.msg_type in HISTORICAL_MESSAGE_TYPES and record.period == period
            ]
            period_results[direction] = select_cluster_from_normalized_records(
                direction, period_records, config, result.fragment_info, parsed_day
            )
        result.period_directions[period] = period_results

    _spool_path, result.spool_sha256 = write_normalized_spool(config, parsed_day, all_rows)

    result.elapsed_seconds = time.perf_counter() - started
    return result


def process_source_file(
    source_file: Path,
    parsed_day: date,
    config: AppConfig,
    callback: ProgressCallback | None = None,
    cancel_event: threading.Event | None = None,
) -> DayResult:
    """Compatibility wrapper for callers that intentionally process one fragment."""
    identity = parse_source_filename(source_file.name)
    fragment = SourceFragment(
        port=identity.port if identity is not None else config.port,
        day=parsed_day,
        suffix=identity.suffix if identity is not None else None,
        path=source_file,
    )
    return process_day_input(
        DayInput(port=fragment.port, day=parsed_day, fragments=(fragment,)),
        config,
        callback=callback,
        cancel_event=cancel_event,
    )


def build_processing_job(config: AppConfig) -> ProcessingJob:
    discovered, warnings = discover_source_files(config.input_dir, config.port)
    days = tuple(
        sorted(
            (day_input for parsed, day_input in discovered.items()
            if parsed.year == config.year and parsed.month == config.month
            ),
            key=lambda item: item.day,
        )
    )
    if not days:
        raise ValueError(
            f"來源資料夾內找不到港別 {config.port}、{config.year} 年 {config.month} 月的 "
            "D&TMOK <PORT> 每日檔案。"
        )
    return ProcessingJob(
        port=config.port,
        year=config.year,
        month=config.month,
        days=days,
        warnings=tuple(warnings),
    )


def process_month(
    config: AppConfig,
    callback: ProgressCallback | None = None,
    cancel_event: threading.Event | None = None,
) -> tuple[list[DayResult], list[str]]:
    config.validate()
    job = build_processing_job(config)
    month_days = {day_input.day: day_input for day_input in job.days}
    warnings = list(job.warnings)

    emit(
        callback,
        kind="job",
        port=job.port,
        year=job.year,
        month=job.month,
        source_count=len(job.days),
        fragment_count=sum(len(day_input.fragments) for day_input in job.days),
    )

    required, free = preflight_output_space(
        config,
        (fragment.path for day_input in job.days for fragment in day_input.fragments),
    )
    emit(callback, kind="preflight", required_bytes=required, free_bytes=free)

    ordered_days = list(job.days)
    if config.max_files is not None:
        ordered_days = ordered_days[: config.max_files]
        allowed_days = {day_input.day for day_input in ordered_days}
    else:
        allowed_days = set(month_days)

    processed: dict[date, DayResult] = {}
    total_files = len(ordered_days)
    pending: list[DayInput] = []
    completed_count = 0
    for day_input in ordered_days:
        cached = load_cached_day(day_input, config)
        if cached is not None:
            processed[day_input.day] = cached
            completed_count += 1
            emit(
                callback,
                kind="cache_hit",
                position=completed_count,
                total=total_files,
                file=f"{day_input.day:%Y-%m-%d}（{len(day_input.fragments)} 個分片）",
            )
        else:
            pending.append(day_input)

    if config.workers == 1 or len(pending) <= 1:
        for day_input in pending:
            if cancel_event and cancel_event.is_set():
                raise CancelledError("使用者已取消")
            emit(
                callback,
                kind="file_start",
                position=completed_count + 1,
                total=total_files,
                file=f"{day_input.day:%Y-%m-%d}（{len(day_input.fragments)} 個分片）",
            )
            result = process_day_input(
                day_input,
                config,
                callback=callback,
                cancel_event=cancel_event,
            )
            processed[day_input.day] = result
            save_cached_day(day_input, config, result)
            completed_count += 1
            emit(
                callback,
                kind="file_done",
                position=completed_count,
                total=total_files,
                file=f"{day_input.day:%Y-%m-%d}（{len(day_input.fragments)} 個分片）",
                seconds=result.elapsed_seconds,
            )
    elif pending:
        emit(callback, kind="parallel_start", workers=config.workers, total=len(pending))
        with concurrent.futures.ProcessPoolExecutor(max_workers=config.workers) as executor:
            futures = {
                executor.submit(_process_day_worker, (day_input, config)): day_input
                for day_input in pending
            }
            try:
                for future in concurrent.futures.as_completed(futures):
                    if cancel_event and cancel_event.is_set():
                        for other in futures:
                            other.cancel()
                        raise CancelledError("使用者已取消")
                    day_input = futures[future]
                    expected_day = day_input.day
                    parsed, result = future.result()
                    if parsed != expected_day:
                        raise RuntimeError(f"平行處理日期不一致：{parsed} / {expected_day}")
                    processed[parsed] = result
                    save_cached_day(day_input, config, result)
                    completed_count += 1
                    emit(
                        callback,
                        kind="file_done",
                        position=completed_count,
                        total=total_files,
                        file=f"{day_input.day:%Y-%m-%d}（{len(day_input.fragments)} 個分片）",
                        seconds=result.elapsed_seconds,
                    )
            except Exception:
                for future in futures:
                    future.cancel()
                raise

    days_in_month = calendar.monthrange(config.year, config.month)[1]
    output: list[DayResult] = []
    for day_number in range(1, days_in_month + 1):
        parsed = date(config.year, config.month, day_number)
        if parsed in processed:
            output.append(processed[parsed])
        elif config.max_files is not None and parsed in month_days and parsed not in allowed_days:
            output.append(empty_day_result(parsed, "測試模式未處理此檔"))
        else:
            output.append(empty_day_result(parsed))
            warnings.append(f"缺少 {parsed:%Y-%m-%d} 的來源檔")
    return output, warnings


def _scope_directions(day_result: DayResult, scope: str) -> dict[str, DirectionResult]:
    if scope == LOGICAL_DAY:
        return day_result.directions
    if scope in (PERIOD_A, PERIOD_B):
        return day_result.period_directions.get(scope, {})
    raise ValueError(f"未知的覆核 scope：{scope}")


def source_manifest_hash(config: AppConfig, day_results: Sequence[DayResult]) -> str:
    payload = {
        "schema": REVIEW_WORKBOOK_SCHEMA,
        "port": config.port,
        "year": config.year,
        "month": config.month,
        "days": [
            {
                "day": result.day.isoformat(),
                "spool_sha256": result.spool_sha256,
                "fragments": [
                    {
                        "sha256": info.sha256,
                        "sheet": info.sheet,
                        "rows": info.rows_scanned,
                    }
                    for info in result.fragment_info
                ],
            }
            for result in day_results
        ],
    }
    encoded = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()


def build_default_decision_snapshot(
    config: AppConfig,
    day_results: Sequence[DayResult],
) -> DecisionSnapshot:
    decisions: list[ReviewDecision] = []
    for day_result in day_results:
        for scope in (LOGICAL_DAY, PERIOD_A, PERIOD_B):
            directions = _scope_directions(day_result, scope)
            for direction_index in OPEN_SEA_INDEXES:
                direction = DIRECTION_ORDER[direction_index]
                result = directions.get(direction, DirectionResult(direction=direction))
                decisions.append(
                    ReviewDecision(
                        scope=scope,
                        day=day_result.day,
                        direction=direction,
                        candidate_id=result.selected_candidate_id,
                    )
                )
    return DecisionSnapshot(
        port=config.port,
        year=config.year,
        month=config.month,
        source_manifest_hash=source_manifest_hash(config, day_results),
        decisions=tuple(decisions),
    )


def _display_candidate_catalog(
    day_results: Sequence[DayResult],
) -> dict[tuple[str, date, str], dict[str, Candidate]]:
    catalog: dict[tuple[str, date, str], dict[str, Candidate]] = {}
    for day_result in day_results:
        for scope in (LOGICAL_DAY, PERIOD_A, PERIOD_B):
            for direction, result in _scope_directions(day_result, scope).items():
                key = (scope, day_result.day, direction)
                catalog[key] = {
                    candidate.candidate_id: candidate
                    for candidate in result.candidates
                    if candidate.candidate_id
                }
    return catalog


def resolve_decision_values(
    snapshot: DecisionSnapshot,
    day_results: Sequence[DayResult],
) -> dict[tuple[str, date, str], float | None]:
    catalog = _display_candidate_catalog(day_results)
    values: dict[tuple[str, date, str], float | None] = {}
    for decision in snapshot.decisions:
        key = (decision.scope, decision.day, decision.direction)
        if decision.scope in (PERIOD_A, PERIOD_B) and decision.forced_numeric is not None:
            raise ValueError(
                f"{decision.day} {decision.scope} {decision.direction} 不允許任意 numeric override。"
            )
        if decision.scope == LOGICAL_DAY and decision.forced_numeric is not None:
            if not math.isfinite(decision.forced_numeric) or decision.forced_numeric < 0:
                raise ValueError(f"{decision.day} {decision.direction} 的 logical-day numeric override 無效。")
            values[key] = decision.forced_numeric
            continue
        if not decision.candidate_id:
            values[key] = None
            continue
        candidate = catalog.get(key, {}).get(decision.candidate_id)
        if candidate is None:
            raise ValueError(
                f"{decision.day} {decision.scope} {decision.direction} 的 candidate ID "
                f"{decision.candidate_id} 不在可驗證候選清單。"
            )
        values[key] = candidate.distance
    return values


def read_review_decisions(
    workbook_path: Path,
    *,
    expected_manifest_hash: str | None = None,
    candidate_catalog: dict[tuple[str, date, str], set[str]] | None = None,
) -> DecisionSnapshot:
    workbook = openpyxl.load_workbook(workbook_path, read_only=False, data_only=False)
    try:
        if "系統資料" not in workbook.sheetnames or "決策台帳" not in workbook.sheetnames:
            raise ValueError("不是 v1.5.0 覆核 workbook：缺少系統資料或決策台帳。")
        system = workbook["系統資料"]
        metadata = {
            str(system.cell(row=row, column=1).value): system.cell(row=row, column=2).value
            for row in range(1, system.max_row + 1)
        }
        if metadata.get("Schema") != REVIEW_WORKBOOK_SCHEMA:
            raise ValueError("覆核 workbook schema 版本不相容。")
        manifest_hash = str(metadata.get("Source manifest SHA-256", ""))
        if expected_manifest_hash is not None and manifest_hash != expected_manifest_hash:
            raise ValueError("來源 fragment manifest 已變更；不可沿用舊覆核決策。")
        port = normalize_port(str(metadata.get("Port", "")))
        period = str(metadata.get("Period", ""))
        match = re.fullmatch(r"(\d{4})-(\d{2})", period)
        if match is None:
            raise ValueError("覆核 workbook 的 Period metadata 無效。")
        year, month = map(int, match.groups())

        ledger = workbook["決策台帳"]
        expected_headers = [
            "Scope",
            "日期",
            "方位",
            "自動 Candidate ID",
            "自動值",
            "決策 Candidate ID",
            "Forced numeric override",
            "Final value",
            "狀態",
            "備註",
        ]
        actual_headers = [ledger.cell(row=1, column=index).value for index in range(1, 11)]
        if actual_headers != expected_headers:
            raise ValueError("決策台帳欄位已被變更，無法安全 finalization。")
        decisions: list[ReviewDecision] = []
        seen: set[tuple[str, date, str]] = set()
        for row in range(2, ledger.max_row + 1):
            scope_value = ledger.cell(row=row, column=1).value
            if scope_value is None:
                continue
            scope = str(scope_value)
            if scope not in (LOGICAL_DAY, PERIOD_A, PERIOD_B):
                raise ValueError(f"決策台帳 A{row} 的 scope 無效：{scope}")
            raw_day = ledger.cell(row=row, column=2).value
            if isinstance(raw_day, datetime):
                parsed_day = raw_day.date()
            elif isinstance(raw_day, date):
                parsed_day = raw_day
            else:
                raise ValueError(f"決策台帳 B{row} 的日期無效。")
            direction = str(ledger.cell(row=row, column=3).value)
            if direction not in OPEN_SEA_DIRECTIONS:
                raise ValueError(f"決策台帳 C{row} 的方位不在 21 個海向內。")
            key = (scope, parsed_day, direction)
            if key in seen:
                raise ValueError(f"決策台帳第 {row} 列與前列重複：{key}")
            seen.add(key)
            raw_candidate = ledger.cell(row=row, column=6).value
            candidate_id = str(raw_candidate).strip() if raw_candidate not in (None, "") else None
            raw_forced = ledger.cell(row=row, column=7).value
            if raw_forced in (None, ""):
                forced_numeric = None
            elif isinstance(raw_forced, (int, float)) and math.isfinite(float(raw_forced)):
                forced_numeric = float(raw_forced)
            else:
                raise ValueError(f"決策台帳 G{row} 的 forced numeric override 無效。")
            if scope in (PERIOD_A, PERIOD_B) and forced_numeric is not None:
                raise ValueError(f"決策台帳 G{row}：Period A/B 不允許 numeric override。")
            if candidate_catalog is not None and candidate_id:
                if candidate_id not in candidate_catalog.get(key, set()):
                    raise ValueError(f"決策台帳 F{row} 的 candidate ID 不屬於該 day/period/direction。")
            final_formula = ledger.cell(row=row, column=8).value
            if not isinstance(final_formula, str) or not final_formula.startswith("="):
                raise ValueError(f"決策台帳 H{row} 的 Final value 公式已遭破壞。")
            note_value = ledger.cell(row=row, column=10).value
            decisions.append(
                ReviewDecision(
                    scope=scope,
                    day=parsed_day,
                    direction=direction,
                    candidate_id=candidate_id,
                    forced_numeric=forced_numeric,
                    note="" if note_value is None else str(note_value),
                )
            )
        return DecisionSnapshot(
            port=port,
            year=year,
            month=month,
            source_manifest_hash=manifest_hash,
            decisions=tuple(decisions),
        )
    finally:
        workbook.close()


def _excel_col(column_zero_based: int) -> str:
    value = column_zero_based + 1
    letters = ""
    while value:
        value, remainder = divmod(value - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _safe_sheet_name(name: str) -> str:
    return re.sub(r"[\[\]:*?/\\]", "_", name)[:31]


def _source_fragment_label(day_result: DayResult) -> str:
    if not day_result.source_fragments:
        return "缺檔"
    return "、".join(path.name for path in day_result.source_fragments)


def _stats(values: Iterable[float | None]) -> tuple[float | None, float | None, float | None, float | None]:
    valid = [float(value) for value in values if value is not None]
    if not valid:
        return None, None, None, None
    stdev = statistics.stdev(valid) if len(valid) >= 2 else None
    return max(valid), min(valid), statistics.mean(valid), stdev


def write_monthly_workbook(
    config: AppConfig,
    day_results: list[DayResult],
    warnings: list[str],
    decision_snapshot: DecisionSnapshot | None = None,
    callback: ProgressCallback | None = None,
    cancel_event: threading.Event | None = None,
) -> Path:
    if decision_snapshot is None:
        decision_snapshot = build_default_decision_snapshot(config, day_results)
    if (
        decision_snapshot.port != config.port
        or decision_snapshot.year != config.year
        or decision_snapshot.month != config.month
    ):
        raise ValueError("決策 snapshot 與目前港別／月份不一致。")
    expected_manifest = source_manifest_hash(config, day_results)
    if decision_snapshot.source_manifest_hash != expected_manifest:
        raise ValueError("決策 snapshot 的來源 fragment manifest 不一致。")
    decision_values = resolve_decision_values(decision_snapshot, day_results)
    decisions_by_key = decision_snapshot.by_key()
    warnings = list(warnings) + [
        f"{day_result.day:%Y-%m-%d}：{diagnostic}"
        for day_result in day_results
        for diagnostic in day_result.diagnostics
    ]
    config.output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = config.output_path.with_name(config.output_path.stem + ".building.xlsx")
    if temp_path.exists():
        temp_path.unlink()

    emit(callback, kind="write_start", file=config.output_path.name)
    workbook = xlsxwriter.Workbook(
        temp_path,
        {
            "nan_inf_to_errors": True,
            "default_date_format": "yyyy-mm-dd",
        },
    )
    workbook.set_properties(
        {
            "title": f"{config.port} {config.year}年{config.month:02d}月 AIS 32方位數值",
            "subject": "AIS 通訊距離自動整理與複核",
            "author": "AIS 32方位月報一鍵製作",
            "comments": f"由原始 D&TMOK {config.port} 每日檔自動產生；原始檔未修改。",
        }
    )
    workbook.set_calc_mode("auto")

    formats = {
        "title": workbook.add_format(
            {"font_name": "Microsoft JhengHei", "font_size": 18, "bold": True, "font_color": "#FFFFFF", "bg_color": "#17324D", "align": "left", "valign": "vcenter"}
        ),
        "section": workbook.add_format(
            {"font_name": "Microsoft JhengHei", "font_size": 12, "bold": True, "font_color": "#17324D", "bg_color": "#DCEAF5", "bottom": 1, "bottom_color": "#8FAFC6"}
        ),
        "header": workbook.add_format(
            {"font_name": "Microsoft JhengHei", "bold": True, "font_color": "#FFFFFF", "bg_color": "#2F6B7C", "align": "center", "valign": "vcenter", "text_wrap": True, "bottom": 1, "bottom_color": "#17324D"}
        ),
        "subheader": workbook.add_format(
            {"font_name": "Microsoft JhengHei", "bold": True, "font_color": "#17324D", "bg_color": "#EAF3F7", "align": "center", "valign": "vcenter", "text_wrap": True}
        ),
        "body": workbook.add_format({"font_name": "Microsoft JhengHei", "font_size": 10, "valign": "top"}),
        "body_wrap": workbook.add_format({"font_name": "Microsoft JhengHei", "font_size": 10, "valign": "top", "text_wrap": True}),
        "number": workbook.add_format({"font_name": "Microsoft JhengHei", "font_size": 10, "num_format": "0.000", "align": "right"}),
        "integer": workbook.add_format({"font_name": "Microsoft JhengHei", "font_size": 10, "num_format": "#,##0", "align": "right"}),
        "date": workbook.add_format({"font_name": "Microsoft JhengHei", "font_size": 10, "num_format": "yyyy-mm-dd", "align": "center"}),
        "timestamp": workbook.add_format({"font_name": "Microsoft JhengHei", "font_size": 10, "num_format": "yyyy-mm-dd hh:mm:ss"}),
        "auto": workbook.add_format({"font_name": "Microsoft JhengHei", "bg_color": "#E6F4EA", "font_color": "#1F6D3A", "num_format": "0.000", "align": "center"}),
        "review": workbook.add_format({"font_name": "Microsoft JhengHei", "bg_color": "#FFF2CC", "font_color": "#8A5A00", "num_format": "0.000", "align": "center"}),
        "missing": workbook.add_format({"font_name": "Microsoft JhengHei", "bg_color": "#FCE8E6", "font_color": "#A61B1B", "align": "center"}),
        "excluded": workbook.add_format({"font_name": "Microsoft JhengHei", "bg_color": "#E7E9EC", "font_color": "#6B7280", "align": "center"}),
        "input": workbook.add_format({"font_name": "Microsoft JhengHei", "bg_color": "#FFF9E6", "font_color": "#7A4B00", "num_format": "0.000", "align": "center", "border": 1, "border_color": "#E0C36C"}),
        "final": workbook.add_format({"font_name": "Microsoft JhengHei", "bold": True, "bg_color": "#DCEAF5", "font_color": "#17324D", "num_format": "0.000", "align": "center"}),
        "note": workbook.add_format({"font_name": "Microsoft JhengHei", "font_size": 9, "font_color": "#5B6573", "italic": True, "text_wrap": True}),
        "stat_label": workbook.add_format({"font_name": "Microsoft JhengHei", "bold": True, "font_color": "#FFFFFF", "bg_color": "#17324D", "align": "center"}),
        "stat": workbook.add_format({"font_name": "Microsoft JhengHei", "bold": True, "num_format": "0.000", "bg_color": "#EDF4F7", "align": "center"}),
        "link": workbook.add_format({"font_name": "Microsoft JhengHei", "font_color": "#0563C1", "underline": True, "align": "center"}),
        "ledger_input": workbook.add_format({"font_name": "Consolas", "font_size": 9, "bg_color": "#FFF9E6", "font_color": "#7A4B00", "locked": False}),
        "ledger_note": workbook.add_format({"font_name": "Microsoft JhengHei", "font_size": 9, "bg_color": "#FFF9E6", "locked": False, "text_wrap": True}),
    }

    guide = workbook.add_worksheet("操作說明")
    guide.hide_gridlines(2)
    guide.set_column("A:A", 22)
    guide.set_column("B:B", 92)
    guide.set_row(0, 34)
    guide.merge_range("A1:B1", APP_TITLE, formats["title"])
    guide.write("A3", "本次設定", formats["section"])
    settings = [
        ("港別", config.port),
        ("製作月份", f"{config.year} 年 {config.month} 月"),
        ("來源資料夾", str(config.input_dir)),
        ("輸出檔", str(config.output_path)),
        ("訊息類型", ", ".join(map(str, config.message_types))),
        ("經度描述", "LONGITUDE_DESC = East"),
        ("距離上限", f"{config.max_distance:g} NM；超過者不納入自動值並保留計數"),
        ("群聚規則", f"由高至低尋找至少 {config.cluster_size} 筆、彼此位於最高值減 {config.tolerance:.0%} 範圍內的第一群"),
        ("候選保留", f"每個方向保留前 {config.top_candidates} 筆合格資料，供追查與人工覆核"),
        ("處理方位", "北至東南東、西南西至北微西，共 21 方位；朝向臺灣的中間 11 方位不處理"),
        ("原始資料", "所有來源檔只讀取、不修改；清理與覆核結果另存於本月報"),
    ]
    for row, (label, value) in enumerate(settings, start=3):
        guide.write(row, 0, label, formats["subheader"])
        guide.write(row, 1, value, formats["body_wrap"])
        guide.set_row(row, 30 if len(value) > 70 else 22)

    next_row = 3 + len(settings) + 2
    guide.write(next_row, 0, "使用方式", formats["section"])
    instructions = [
        "先看「總表」與兩張圖，這些值會引用各日工作表的「最終值」。",
        "再看「待複核」；黃色項目代表群聚不足，或西南西／西微南／西超過 10 NM。",
        "如要修正，前往該日工作表，在方向欄的「人工覆核值」輸入距離；「最終值」與總表會自動改用人工值。",
        "每日工作表 A:F 保存自動判斷所用的候選資料與原始列號，可回查來源檔。",
        "自動規則是把影片中的人工判斷具體化；它不能取代研究上的最終判斷，因此所有例外均明確標記。",
    ]
    for offset, instruction in enumerate(instructions, start=1):
        guide.write(next_row + offset, 0, f"{offset}.", formats["subheader"])
        guide.write(next_row + offset, 1, instruction, formats["body_wrap"])
        guide.set_row(next_row + offset, 32)

    warning_row = next_row + len(instructions) + 2
    guide.write(warning_row, 0, "檔案警告", formats["section"])
    if warnings:
        for offset, warning in enumerate(warnings, start=1):
            guide.write(warning_row + offset, 0, offset, formats["integer"])
            guide.write(warning_row + offset, 1, warning, formats["body_wrap"])
    else:
        guide.write(warning_row + 1, 1, "無", formats["body"])
    guide.freeze_panes(2, 0)
    guide.set_landscape()
    guide.fit_to_pages(1, 0)

    candidate_sheet = workbook.add_worksheet("候選清單")
    candidate_sheet.hide_gridlines(2)
    candidate_sheet.freeze_panes(1, 0)
    candidate_headers = [
        "Scope",
        "日期",
        "方位",
        "排名",
        "Candidate ID",
        "Distance (NM)",
        "Bearing",
        "Timestamp",
        "MMSI",
        "Channel",
        "Fragment",
        "Sheet",
        "Source row",
        "Msg type",
    ]
    candidate_sheet.write_row(0, 0, candidate_headers, formats["header"])
    candidate_sheet.set_column("A:A", 14)
    candidate_sheet.set_column("B:B", 12)
    candidate_sheet.set_column("C:D", 11)
    candidate_sheet.set_column("E:E", 27)
    candidate_sheet.set_column("F:G", 13)
    candidate_sheet.set_column("H:H", 21)
    candidate_sheet.set_column("I:J", 13)
    candidate_sheet.set_column("K:K", 38)
    candidate_sheet.set_column("L:N", 13)
    candidate_row = 1
    for day_result in day_results:
        for scope in (LOGICAL_DAY, PERIOD_A, PERIOD_B):
            for direction_index in OPEN_SEA_INDEXES:
                direction = DIRECTION_ORDER[direction_index]
                direction_result = _scope_directions(day_result, scope).get(direction)
                if direction_result is None:
                    continue
                for list_position, candidate in enumerate(direction_result.candidates, start=1):
                    candidate_sheet.write(candidate_row, 0, scope, formats["body"])
                    candidate_sheet.write_datetime(
                        candidate_row,
                        1,
                        datetime.combine(day_result.day, datetime.min.time()),
                        formats["date"],
                    )
                    candidate_sheet.write(candidate_row, 2, direction, formats["body"])
                    candidate_sheet.write_number(
                        candidate_row, 3, candidate.rank or list_position, formats["integer"]
                    )
                    candidate_sheet.write(candidate_row, 4, candidate.candidate_id, formats["body"])
                    candidate_sheet.write_number(candidate_row, 5, candidate.distance, formats["number"])
                    candidate_sheet.write_number(candidate_row, 6, candidate.bearing, formats["number"])
                    if candidate.timestamp is None:
                        candidate_sheet.write_blank(candidate_row, 7, None, formats["body"])
                    else:
                        candidate_sheet.write_datetime(
                            candidate_row,
                            7,
                            candidate.timestamp,
                            formats["timestamp"],
                        )
                    if candidate.mmsi is None:
                        candidate_sheet.write_blank(candidate_row, 8, None, formats["body"])
                    else:
                        candidate_sheet.write_number(candidate_row, 8, candidate.mmsi, formats["integer"])
                    candidate_sheet.write(candidate_row, 9, candidate.channel or "", formats["body"])
                    candidate_sheet.write(candidate_row, 10, candidate.fragment, formats["body"])
                    candidate_sheet.write(candidate_row, 11, candidate.sheet, formats["body"])
                    candidate_sheet.write_number(candidate_row, 12, candidate.source_row, formats["integer"])
                    if candidate.msg_type is None:
                        candidate_sheet.write_blank(candidate_row, 13, None, formats["body"])
                    else:
                        candidate_sheet.write_number(candidate_row, 13, candidate.msg_type, formats["integer"])
                    candidate_row += 1
    if candidate_row > 1:
        candidate_sheet.autofilter(0, 0, candidate_row - 1, len(candidate_headers) - 1)

    ledger = workbook.add_worksheet("決策台帳")
    ledger.hide_gridlines(2)
    ledger.freeze_panes(1, 0)
    ledger_headers = [
        "Scope",
        "日期",
        "方位",
        "自動 Candidate ID",
        "自動值",
        "決策 Candidate ID",
        "Forced numeric override",
        "Final value",
        "狀態",
        "備註",
    ]
    ledger.write_row(0, 0, ledger_headers, formats["header"])
    ledger.set_column("A:A", 14)
    ledger.set_column("B:B", 12)
    ledger.set_column("C:C", 11)
    ledger.set_column("D:D", 27)
    ledger.set_column("E:E", 13)
    ledger.set_column("F:F", 27)
    ledger.set_column("G:H", 18)
    ledger.set_column("I:I", 13)
    ledger.set_column("J:J", 42)
    decision_row_lookup: dict[tuple[str, date, str], int] = {}
    candidate_last_excel_row = max(candidate_row, 2)
    lookup_formula = (
        f"IFERROR(INDEX('候選清單'!$F$2:$F${candidate_last_excel_row},"
        f"MATCH(F{{row}},'候選清單'!$E$2:$E${candidate_last_excel_row},0)),\"#INVALID_CANDIDATE\")"
    )
    for ledger_row, decision in enumerate(decision_snapshot.decisions, start=1):
        key = (decision.scope, decision.day, decision.direction)
        decision_row_lookup[key] = ledger_row + 1
        automatic = _scope_directions(
            next(item for item in day_results if item.day == decision.day), decision.scope
        )[decision.direction]
        ledger.write(ledger_row, 0, decision.scope, formats["body"])
        ledger.write_datetime(
            ledger_row,
            1,
            datetime.combine(decision.day, datetime.min.time()),
            formats["date"],
        )
        ledger.write(ledger_row, 2, decision.direction, formats["body"])
        ledger.write(ledger_row, 3, automatic.selected_candidate_id or "", formats["body"])
        if automatic.selected is None:
            ledger.write_blank(ledger_row, 4, None, formats["missing"])
        else:
            ledger.write_number(ledger_row, 4, automatic.selected, formats["number"])
        ledger.write(ledger_row, 5, decision.candidate_id or "", formats["ledger_input"])
        if decision.forced_numeric is None:
            ledger.write_blank(
                ledger_row,
                6,
                None,
                formats["ledger_input"] if decision.scope == LOGICAL_DAY else formats["excluded"],
            )
        else:
            ledger.write_number(ledger_row, 6, decision.forced_numeric, formats["ledger_input"])
        excel_row = ledger_row + 1
        candidate_lookup = lookup_formula.format(row=excel_row)
        if decision.scope == LOGICAL_DAY:
            formula = (
                f'=IF(ISNUMBER(G{excel_row}),G{excel_row},IF(F{excel_row}="","",{candidate_lookup}))'
            )
        else:
            formula = (
                f'=IF(G{excel_row}<>"","#INVALID_PERIOD_NUMERIC",'
                f'IF(F{excel_row}="","",{candidate_lookup}))'
            )
        final_value = decision_values[key]
        ledger.write_formula(
            ledger_row,
            7,
            formula,
            formats["final"],
            final_value if final_value is not None else "",
        )
        ledger.write(ledger_row, 8, automatic.status, formats["body"])
        ledger.write(ledger_row, 9, decision.note, formats["ledger_note"])
    if decision_snapshot.decisions:
        ledger.autofilter(0, 0, len(decision_snapshot.decisions), len(ledger_headers) - 1)
    ledger.protect("", {"select_unlocked_cells": True, "autofilter": True})

    system_sheet = workbook.add_worksheet("系統資料")
    system_rows = [
        ("Schema", REVIEW_WORKBOOK_SCHEMA),
        ("Port", config.port),
        ("Period", f"{config.year}-{config.month:02d}"),
        ("Source manifest SHA-256", decision_snapshot.source_manifest_hash),
        ("Modern message types", ",".join(map(str, config.message_types))),
        ("Historical delivery message types", ",".join(map(str, HISTORICAL_MESSAGE_TYPES))),
        ("Decision count", len(decision_snapshot.decisions)),
    ]
    for system_row, (label, value) in enumerate(system_rows):
        system_sheet.write(system_row, 0, label)
        system_sheet.write(system_row, 1, value)
    manifest_row = len(system_rows) + 1
    system_sheet.write_row(
        manifest_row,
        0,
        ["Date", "Spool SHA-256", "Fragment SHA-256", "Fragment", "Sheet", "Rows", "Aliases"],
    )
    manifest_row += 1
    for day_result in day_results:
        for info in day_result.fragment_info:
            system_sheet.write(manifest_row, 0, day_result.day.isoformat())
            system_sheet.write(manifest_row, 1, day_result.spool_sha256)
            system_sheet.write(manifest_row, 2, info.sha256)
            system_sheet.write(manifest_row, 3, info.path.name)
            system_sheet.write(manifest_row, 4, info.sheet)
            system_sheet.write_number(manifest_row, 5, info.rows_scanned)
            system_sheet.write(manifest_row, 6, "、".join(path.name for path in info.aliases))
            manifest_row += 1
    system_sheet.hide()

    daily_sheet_names: dict[date, str] = {}
    for day_index, day_result in enumerate(day_results, start=1):
        if cancel_event and cancel_event.is_set():
            workbook.close()
            raise CancelledError("使用者已取消")
        sheet_name = _safe_sheet_name(f"{config.month}月{day_result.day.day}日")
        daily_sheet_names[day_result.day] = sheet_name
        sheet = workbook.add_worksheet(sheet_name)
        sheet.hide_gridlines(2)
        sheet.freeze_panes(1, 0)
        sheet.set_column("A:A", 11)
        sheet.set_column("B:B", 11)
        sheet.set_column("C:C", 11)
        sheet.set_column("D:D", 12)
        sheet.set_column("E:E", 11)
        sheet.set_column("F:F", 38)
        sheet.set_column("G:G", 13)
        sheet.set_column("H:AM", 10)
        sheet.set_row(0, 32)
        candidate_headers = ["DISTANCE", "DEGREE", "方位", "候選排名", "原始列", "判定說明"]
        for column, header in enumerate(candidate_headers):
            sheet.write(0, column, header, formats["header"])

        for direction_index, direction in enumerate(DIRECTION_ORDER):
            column = 7 + direction_index
            if direction_index in OPEN_SEA_INDEXES:
                sheet.write(0, column, direction, formats["header"])
            else:
                sheet.write(0, column, direction, formats["excluded"])
        sheet.write(1, 6, "自動值", formats["subheader"])
        sheet.write(2, 6, "覆核請至決策台帳", formats["subheader"])
        sheet.write(3, 6, "Logical-day final", formats["subheader"])
        sheet.write(4, 6, "判定", formats["subheader"])
        sheet.write(5, 6, "來源檔", formats["subheader"])

        candidate_rows: list[tuple[float, float, str, int, int, str]] = []
        for direction_index, direction in enumerate(DIRECTION_ORDER):
            column = 7 + direction_index
            if direction_index not in OPEN_SEA_INDEXES:
                for row in range(1, 6):
                    sheet.write(row, column, "不處理" if row == 4 else "", formats["excluded"])
                continue
            direction_result = day_result.directions[direction]
            selected = direction_result.selected
            status_format = formats["auto"] if direction_result.status == "自動採用" else formats["review"] if direction_result.status == "待複核" else formats["missing"]
            if selected is None:
                sheet.write_blank(1, column, None, status_format)
            else:
                sheet.write_number(1, column, selected, status_format)
            ledger_excel_row = decision_row_lookup[(LOGICAL_DAY, day_result.day, direction)]
            sheet.write_url(
                2,
                column,
                f"internal:'決策台帳'!F{ledger_excel_row}",
                formats["link"],
                "前往台帳",
            )
            final_value = decision_values[(LOGICAL_DAY, day_result.day, direction)]
            final_formula = f"='決策台帳'!H{ledger_excel_row}"
            sheet.write_formula(
                3,
                column,
                final_formula,
                formats["final"],
                final_value if final_value is not None else "",
            )
            sheet.write(4, column, direction_result.status, status_format)
            sheet.write(5, column, _source_fragment_label(day_result), formats["note"])

            for list_position, candidate in enumerate(direction_result.candidates, start=1):
                rank = candidate.rank or list_position
                note = ""
                if direction_result.selected_rank is not None:
                    if rank < direction_result.selected_rank:
                        note = "孤立高值，不採用"
                    elif direction_result.selected_rank <= rank < direction_result.selected_rank + direction_result.cluster_count:
                        note = "採用群聚"
                candidate_rows.append((candidate.distance, candidate.bearing, direction, rank, candidate.source_row, note))

        for row, values in enumerate(candidate_rows, start=1):
            sheet.write_number(row, 0, values[0], formats["number"])
            sheet.write_number(row, 1, values[1], formats["number"])
            sheet.write(row, 2, values[2], formats["body"])
            sheet.write_number(row, 3, values[3], formats["integer"])
            sheet.write_number(row, 4, values[4], formats["integer"])
            sheet.write(row, 5, values[5], formats["body"])
        if candidate_rows:
            sheet.autofilter(0, 0, len(candidate_rows), 5)
        sheet.set_landscape()
        sheet.fit_to_pages(1, 0)
        emit(callback, kind="sheet_written", position=day_index, total=len(day_results), sheet=sheet_name)

    total_sheet = workbook.add_worksheet("總表")
    total_sheet.hide_gridlines(2)
    total_sheet.freeze_panes(4, 1)
    total_sheet.set_column("A:A", 13)
    total_sheet.set_column("B:AG", 10)
    total_sheet.set_column("AI:AL", 13)
    total_sheet.set_row(0, 34)
    total_sheet.merge_range("A1:AG1", f"{config.port} {config.year} 年 {config.month:02d} 月 AIS 32方位通訊距離", formats["title"])
    total_sheet.merge_range("A2:AG2", f"自動規則：{config.max_distance:g} NM 上限；{config.tolerance:.0%} 群聚；至少 {config.cluster_size} 筆。黃色／紅色儲存格請查看「待複核」。", formats["note"])
    total_sheet.write(3, 0, "日期", formats["header"])
    for direction_index, direction in enumerate(DIRECTION_ORDER):
        column = 1 + direction_index
        total_sheet.write(3, column, direction, formats["header"] if direction_index in OPEN_SEA_INDEXES else formats["excluded"])

    first_data_row_excel = 5
    last_data_row_excel = 4 + len(day_results)
    for row_index, day_result in enumerate(day_results, start=4):
        total_sheet.write_datetime(row_index, 0, datetime.combine(day_result.day, datetime.min.time()), formats["date"])
        source_sheet = daily_sheet_names[day_result.day]
        for direction_index, direction in enumerate(DIRECTION_ORDER):
            column = 1 + direction_index
            if direction_index not in OPEN_SEA_INDEXES:
                total_sheet.write_blank(row_index, column, None, formats["excluded"])
                continue
            direction_result = day_result.directions[direction]
            source_column = _excel_col(7 + direction_index)
            formula = f'=IFERROR(\'{source_sheet}\'!{source_column}4,"")'
            final_value = decision_values[(LOGICAL_DAY, day_result.day, direction)]
            cached = final_value if final_value is not None else ""
            status_format = formats["auto"] if direction_result.status == "自動採用" else formats["review"] if direction_result.status == "待複核" else formats["missing"]
            total_sheet.write_formula(row_index, column, formula, status_format, cached)

    stat_start_row = 5 + len(day_results)
    stat_specs = [("MAX", "MAX"), ("MIN", "MIN"), ("平均", "AVERAGE"), ("標準差", "STDEV.S")]
    stat_values_by_direction: dict[str, tuple[float | None, float | None, float | None, float | None]] = {}
    for direction in DIRECTION_ORDER:
        stat_values_by_direction[direction] = _stats(
            decision_values.get((LOGICAL_DAY, day.day, direction)) for day in day_results
        )
    for stat_offset, (label, function) in enumerate(stat_specs):
        row = stat_start_row + stat_offset
        total_sheet.write(row, 0, label, formats["stat_label"])
        for direction_index, direction in enumerate(DIRECTION_ORDER):
            column = 1 + direction_index
            if direction_index not in OPEN_SEA_INDEXES:
                total_sheet.write_blank(row, column, None, formats["excluded"])
                continue
            col_letter = _excel_col(column)
            formula = f'=IFERROR({function}({col_letter}{first_data_row_excel}:{col_letter}{last_data_row_excel}),"")'
            cached = stat_values_by_direction[direction][stat_offset]
            total_sheet.write_formula(row, column, formula, formats["stat"], cached if cached is not None else "")

    helper_row = 3
    total_sheet.write(helper_row, 34, "海向方位", formats["header"])
    total_sheet.write(helper_row, 35, "MAX", formats["header"])
    total_sheet.write(helper_row, 36, "MIN", formats["header"])
    total_sheet.write(helper_row, 37, "平均", formats["header"])
    for helper_offset, direction_index in enumerate(OPEN_SEA_INDEXES, start=1):
        row = helper_row + helper_offset
        direction = DIRECTION_ORDER[direction_index]
        source_col = _excel_col(1 + direction_index)
        total_sheet.write(row, 34, direction, formats["body"])
        for stat_offset, helper_col in enumerate((35, 36, 37)):
            source_row_excel = stat_start_row + stat_offset + 1
            formula = f'=IFERROR({source_col}{source_row_excel},"")'
            cached = stat_values_by_direction[direction][stat_offset]
            total_sheet.write_formula(row, helper_col, formula, formats["number"], cached if cached is not None else "")

    radar = workbook.add_chart({"type": "radar", "subtype": "with_markers"})
    palette = ["#2F6B7C", "#D97706", "#7C3AED", "#15803D", "#B91C1C", "#0369A1", "#A21CAF"]
    for day_offset, day_result in enumerate(day_results):
        if not any(
            decision_values.get((LOGICAL_DAY, day_result.day, DIRECTION_ORDER[index])) is not None
            for index in OPEN_SEA_INDEXES
        ):
            continue
        excel_row = first_data_row_excel + day_offset
        radar.add_series(
            {
                "name": f"{config.month}/{day_result.day.day}",
                "categories": ["總表", 3, 1, 3, 32],
                "values": ["總表", excel_row - 1, 1, excel_row - 1, 32],
                "line": {"color": palette[day_offset % len(palette)], "width": 0.75, "transparency": 45},
                "marker": {"type": "none"},
            }
        )
    radar.set_title({"name": f"{config.year}-{config.month:02d} AIS 通訊涵蓋圖"})
    radar.set_legend({"position": "bottom", "font": {"size": 8}})
    radar.set_size({"width": 1080, "height": 560})
    radar.set_style(10)

    columns = workbook.add_chart({"type": "column"})
    category_range = ["總表", helper_row + 1, 34, helper_row + len(OPEN_SEA_INDEXES), 34]
    for series_name, helper_col, color in (("最遠", 35, "#2F6B7C"), ("最近", 36, "#D97706")):
        columns.add_series(
            {
                "name": series_name,
                "categories": category_range,
                "values": ["總表", helper_row + 1, helper_col, helper_row + len(OPEN_SEA_INDEXES), helper_col],
                "fill": {"color": color, "transparency": 10},
                "border": {"none": True},
            }
        )
    line = workbook.add_chart({"type": "line"})
    line.add_series(
        {
            "name": "平均",
            "categories": category_range,
            "values": ["總表", helper_row + 1, 37, helper_row + len(OPEN_SEA_INDEXES), 37],
            "line": {"color": "#B91C1C", "width": 2.25},
            "marker": {"type": "circle", "size": 4, "border": {"color": "#B91C1C"}, "fill": {"color": "#FFFFFF"}},
        }
    )
    columns.combine(line)
    columns.set_title({"name": f"{config.year}-{config.month:02d} AIS 通訊最近、最遠與平均距離（NM）"})
    columns.set_y_axis({"name": "海里（NM）", "major_gridlines": {"visible": True, "line": {"color": "#D9E2E8"}}, "num_format": "0"})
    columns.set_x_axis({"label_position": "low", "num_font": {"rotation": -45, "size": 8}})
    columns.set_legend({"position": "bottom"})
    columns.set_size({"width": 1080, "height": 520})
    columns.set_style(10)

    chart_row = stat_start_row + len(stat_specs) + 2
    total_sheet.insert_chart(chart_row, 0, radar)
    total_sheet.insert_chart(chart_row + 29, 0, columns)
    total_sheet.set_landscape()
    total_sheet.fit_to_pages(1, 0)
    total_sheet.print_area(0, 0, chart_row + 55, 32)

    review_sheet = workbook.add_worksheet("待複核")
    review_sheet.hide_gridlines(2)
    review_sheet.freeze_panes(1, 0)
    review_sheet.set_column("A:A", 12)
    review_sheet.set_column("B:B", 11)
    review_sheet.set_column("C:C", 12)
    review_sheet.set_column("D:D", 12)
    review_sheet.set_column("E:E", 12)
    review_sheet.set_column("F:F", 54)
    review_sheet.set_column("G:G", 54)
    review_sheet.set_column("H:H", 22)
    review_headers = ["日期", "方位", "自動值", "群聚筆數", "狀態", "原因", "前10筆候選距離", "人工覆核位置"]
    for column, header in enumerate(review_headers):
        review_sheet.write(0, column, header, formats["header"])
    review_row = 1
    for day_result in day_results:
        for direction_index in OPEN_SEA_INDEXES:
            direction = DIRECTION_ORDER[direction_index]
            direction_result = day_result.directions[direction]
            if direction_result.status == "自動採用":
                continue
            review_sheet.write_datetime(review_row, 0, datetime.combine(day_result.day, datetime.min.time()), formats["date"])
            review_sheet.write(review_row, 1, direction, formats["body"])
            if direction_result.selected is None:
                review_sheet.write_blank(review_row, 2, None, formats["missing"])
            else:
                review_sheet.write_number(review_row, 2, direction_result.selected, formats["review"])
            review_sheet.write_number(review_row, 3, direction_result.cluster_count, formats["integer"])
            review_sheet.write(review_row, 4, direction_result.status, formats["review"] if direction_result.status == "待複核" else formats["missing"])
            review_sheet.write(review_row, 5, direction_result.reason, formats["body_wrap"])
            review_sheet.write(review_row, 6, ", ".join(f"{candidate.distance:.3f}" for candidate in direction_result.candidates[:10]), formats["body_wrap"])
            ledger_excel_row = decision_row_lookup[(LOGICAL_DAY, day_result.day, direction)]
            review_sheet.write_url(
                review_row,
                7,
                f"internal:'決策台帳'!F{ledger_excel_row}",
                formats["link"],
                f"決策台帳!F{ledger_excel_row}",
            )
            review_row += 1
    if review_row == 1:
        review_sheet.write(1, 0, "沒有待複核項目", formats["auto"])
    else:
        review_sheet.autofilter(0, 0, review_row - 1, len(review_headers) - 1)

    log_sheet = workbook.add_worksheet("處理紀錄")
    log_sheet.hide_gridlines(2)
    log_sheet.freeze_panes(5, 0)
    log_sheet.set_column("A:A", 12)
    log_sheet.set_column("B:B", 38)
    log_sheet.set_column("C:G", 14)
    log_sheet.set_column("H:H", 12)
    source_count = sum(len(day_result.source_fragments) for day_result in day_results)
    for row, (label, value) in enumerate(
        (
            ("Port", config.port),
            ("Period", f"{config.year}-{config.month:02d}"),
            ("Source files", source_count),
        )
    ):
        log_sheet.write(row, 0, label, formats["subheader"])
        log_sheet.write(row, 1, value, formats["body"])
    log_headers = ["日期", "來源檔", "掃描列數", "合格列數", "錯誤列數", "訊息類型略過", "非 East 略過", "秒數"]
    for column, header in enumerate(log_headers):
        log_sheet.write(4, column, header, formats["header"])
    for row, day_result in enumerate(day_results, start=5):
        log_sheet.write_datetime(row, 0, datetime.combine(day_result.day, datetime.min.time()), formats["date"])
        log_sheet.write(row, 1, _source_fragment_label(day_result), formats["body"])
        for column, value in enumerate(
            (
                day_result.rows_scanned,
                day_result.rows_accepted,
                day_result.rows_invalid,
                day_result.rows_wrong_message,
                day_result.rows_not_east,
            ),
            start=2,
        ):
            log_sheet.write_number(row, column, value, formats["integer"])
        log_sheet.write_number(row, 7, day_result.elapsed_seconds, formats["number"])
    log_sheet.autofilter(4, 0, 4 + len(day_results), len(log_headers) - 1)

    try:
        workbook.close()
        os.replace(temp_path, config.output_path)
    except Exception:
        try:
            workbook.close()
        except Exception:
            pass
        raise
    emit(callback, kind="write_done", file=config.output_path.name)
    return config.output_path


def write_legacy_workbook(
    config: AppConfig,
    day_results: list[DayResult],
    callback: ProgressCallback | None = None,
    cancel_event: threading.Event | None = None,
) -> Path | None:
    output_path = config.legacy_output_path
    if output_path is None:
        return None
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_name(output_path.stem + ".building.xlsx")
    if temp_path.exists():
        temp_path.unlink()

    emit(callback, kind="legacy_write_start", file=output_path.name)
    workbook = xlsxwriter.Workbook(
        temp_path,
        {
            "constant_memory": True,
            "nan_inf_to_errors": True,
            "default_date_format": "m月d日",
        },
    )
    workbook.use_zip64()
    workbook.set_properties(
        {
            "title": f"{config.port} {config.year}年{config.month:02d}月 32方位數值 原格式相容版",
            "subject": "完整排序資料、32方位結果與舊式總表",
            "author": "AIS 32方位月報一鍵製作",
            "comments": "A:C 保留舊流程的完整排序資料；H:AM 第 2 列為數值化規則自動結果。",
        }
    )
    workbook.set_calc_mode("auto")
    header_format = workbook.add_format({"bold": True, "align": "center", "bottom": 1})
    number_format = workbook.add_format({"num_format": "0.000"})
    date_format = workbook.add_format({"num_format": "m月d日"})
    stat_label_format = workbook.add_format({"bold": True})

    daily_names: dict[date, str] = {}
    try:
        for position, day_result in enumerate(day_results, start=1):
            if cancel_event and cancel_event.is_set():
                raise CancelledError("使用者已取消")
            sheet_name = _safe_sheet_name(f"{config.month}月{day_result.day.day}日")
            daily_names[day_result.day] = sheet_name
            sheet = workbook.add_worksheet(sheet_name)
            sheet.freeze_panes(1, 0)
            sheet.set_column("A:A", 14)
            sheet.set_column("B:B", 12)
            sheet.set_column("C:C", 11)
            sheet.set_column("D:G", 3)
            sheet.set_column("H:AM", 11)
            sheet.write_row(0, 0, ["DISTANCE", "DEGREE", "方位"], header_format)
            sheet.write_row(0, 7, DIRECTION_ORDER, header_format)
            for direction_index, direction in enumerate(DIRECTION_ORDER):
                selected = day_result.directions[direction].selected
                if direction_index in OPEN_SEA_INDEXES and selected is not None:
                    sheet.write_number(1, 7 + direction_index, selected, number_format)

            spool_path = _legacy_spool_path(config, day_result.day)
            if day_result.source_fragments:
                if not spool_path.is_file():
                    raise FileNotFoundError(
                        f"缺少 {day_result.day:%Y-%m-%d} 的原格式資料暫存檔；請重新執行該月份。"
                    )
                counts, records = read_legacy_spool(spool_path)
                if sum(counts) != day_result.rows_normalized:
                    raise ValueError(
                        f"{day_result.day:%Y-%m-%d} 原格式資料筆數不一致："
                        f"暫存 {sum(counts):,} / 紀錄 {day_result.rows_normalized:,}。"
                    )
                excel_row = 1
                for direction_index, distance, bearing in records:
                    if cancel_event and excel_row % 10_000 == 0 and cancel_event.is_set():
                        raise CancelledError("使用者已取消")
                    sheet.write_number(excel_row, 0, distance)
                    sheet.write_number(excel_row, 1, bearing)
                    sheet.write(excel_row, 2, DIRECTION_ORDER[direction_index])
                    excel_row += 1
            emit(
                callback,
                kind="legacy_sheet_written",
                position=position,
                total=len(day_results),
                sheet=sheet_name,
                rows=day_result.rows_normalized,
            )

        summary = workbook.add_worksheet("總表")
        summary.freeze_panes(1, 1)
        summary.set_column("A:A", 11)
        summary.set_column("B:AG", 10)
        summary.write_row(0, 1, LEGACY_DIRECTION_ABBREVIATIONS, header_format)
        for row, day_result in enumerate(day_results, start=1):
            summary.write_datetime(row, 0, datetime.combine(day_result.day, datetime.min.time()), date_format)
            source_sheet = daily_names[day_result.day]
            for direction_index, direction in enumerate(DIRECTION_ORDER):
                if direction_index not in OPEN_SEA_INDEXES:
                    summary.write_blank(row, 1 + direction_index, None, number_format)
                    continue
                source_column = _excel_col(7 + direction_index)
                selected = day_result.directions[direction].selected
                formula = f"=IFERROR('{source_sheet}'!{source_column}2,\"\")"
                summary.write_formula(row, 1 + direction_index, formula, number_format, selected if selected is not None else "")

        stat_start = len(day_results) + 2
        stat_specs = [("MAX", "MAX", 0), ("MIN", "MIN", 1), ("平均值", "AVERAGE", 2)]
        stats_by_direction = {
            direction: (
                _stats(day.directions[direction].selected for day in day_results)
                if direction_index in OPEN_SEA_INDEXES
                else (None, None, None, None)
            )
            for direction_index, direction in enumerate(DIRECTION_ORDER)
        }
        for offset, (label, _function, stat_index) in enumerate(stat_specs):
            row = stat_start + offset
            summary.write(row, 0, label, stat_label_format)
            for direction_index, direction in enumerate(DIRECTION_ORDER):
                column = 1 + direction_index
                if direction_index not in OPEN_SEA_INDEXES:
                    summary.write_blank(row, column, None, number_format)
                    continue
                cached = stats_by_direction[direction][stat_index]
                if cached is None:
                    summary.write_blank(row, column, None, number_format)
                else:
                    summary.write_number(row, column, cached, number_format)

        radar = workbook.add_chart({"type": "radar"})
        palette = ["#4472C4", "#ED7D31", "#70AD47", "#A5A5A5", "#FFC000", "#5B9BD5"]
        for day_offset, day_result in enumerate(day_results):
            if not any(
                day_result.directions[DIRECTION_ORDER[index]].selected is not None
                for index in OPEN_SEA_INDEXES
            ):
                continue
            row = 1 + day_offset
            radar.add_series(
                {
                    "name": f"{config.month}月{day_result.day.day}日",
                    "categories": ["總表", 0, 1, 0, 32],
                    "values": ["總表", row, 1, row, 32],
                    "line": {"color": palette[day_offset % len(palette)], "width": 0.75, "transparency": 35},
                }
            )
        radar.set_title({"name": f"{config.year}-{config.month:02d} AIS通訊涵蓋圖"})
        radar.set_legend({"position": "right", "font": {"size": 8}})
        radar.set_size({"width": 760, "height": 480})
        summary.insert_chart(stat_start + 6, 0, radar)

        columns = workbook.add_chart({"type": "column"})
        categories = ["總表", 0, 1, 0, 32]
        for label, row, color in (
            ("最遠距離", stat_start, "#4472C4"),
            ("最近距離", stat_start + 1, "#ED7D31"),
            ("平均值", stat_start + 2, "#70AD47"),
        ):
            columns.add_series(
                {
                    "name": label,
                    "categories": categories,
                    "values": ["總表", row, 1, row, 32],
                    "fill": {"color": color},
                    "border": {"none": True},
                }
            )
        columns.set_title({"name": f"{config.year}-{config.month:02d} AIS通訊最遠、最近與平均距離（NM）"})
        columns.set_legend({"position": "bottom"})
        columns.set_y_axis({"name": "NM", "num_format": "0"})
        columns.set_size({"width": 760, "height": 480})
        summary.insert_chart(stat_start + 31, 0, columns)

        mapping = workbook.add_worksheet("工作")
        mapping.write_row(0, 9, ["POSITION", "方位"], header_format)
        for degree in range(361):
            direction_info = degree_to_direction_index(degree)
            direction = DIRECTION_ORDER[direction_info[0]] if direction_info else ""
            mapping.write_number(degree + 1, 9, degree)
            mapping.write(degree + 1, 10, direction)
        mapping.hide()

        workbook.close()
        os.replace(temp_path, output_path)
    except Exception:
        try:
            workbook.close()
        except Exception:
            pass
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass
        raise
    emit(callback, kind="legacy_write_done", file=output_path.name)
    return output_path


def run_pipeline(
    config: AppConfig,
    callback: ProgressCallback | None = None,
    cancel_event: threading.Event | None = None,
) -> Path:
    started = time.perf_counter()
    day_results, warnings = process_month(config, callback=callback, cancel_event=cancel_event)
    output = write_monthly_workbook(config, day_results, warnings, callback=callback, cancel_event=cancel_event)
    legacy_output = write_legacy_workbook(config, day_results, callback=callback, cancel_event=cancel_event)
    files = [str(output)] + ([str(legacy_output)] if legacy_output is not None else [])
    emit(callback, kind="complete", file=str(output), files=files, seconds=time.perf_counter() - started)
    return output


def parse_message_types(text: str) -> tuple[int, ...]:
    values: list[int] = []
    for item in re.split(r"[,，\s]+", text.strip()):
        if not item:
            continue
        value = int(item)
        if value not in values:
            values.append(value)
    return tuple(values)


def format_seconds(seconds: float) -> str:
    seconds_int = max(int(round(seconds)), 0)
    hours, remainder = divmod(seconds_int, 3600)
    minutes, secs = divmod(remainder, 60)
    if hours:
        return f"{hours} 小時 {minutes} 分 {secs} 秒"
    if minutes:
        return f"{minutes} 分 {secs} 秒"
    return f"{secs} 秒"


def default_output_filename(port: str, year: int, month: int) -> str:
    return f"{normalize_port(port)}_{year}年{month:02d}月_32方位數值_新版自動分析.xlsx"


def derive_legacy_output_path(modern_output: Path) -> Path:
    stem = modern_output.stem
    for suffix in ("_新版自動分析", "_自動分析版", "_新版"):
        if stem.endswith(suffix):
            stem = stem[: -len(suffix)]
            break
    return modern_output.with_name(f"{stem}_原格式相容版.xlsx")


def friendly_error_message(error: BaseException) -> str:
    if isinstance(error, CancelledError):
        return "已依使用者要求安全停止；正式輸出檔不會被半成品覆寫。"
    if isinstance(error, SourceFileError):
        return str(error)
    if isinstance(error, FileExistsError):
        return f"輸出檔已存在。請改檔名、移走舊檔，或確認允許覆寫。\n\n{error}"
    if isinstance(error, PermissionError):
        return (
            "無法讀寫檔案。最常見原因是輸出 Excel 正在被其他人開啟，或資料夾沒有寫入權限。"
            "請關閉相關 Excel、確認共用磁碟連線後再試。\n\n"
            f"原始錯誤：{error}"
        )
    if isinstance(error, (zipfile.BadZipFile, openpyxl.utils.exceptions.InvalidFileException)):
        return f"其中一個來源檔不是有效的 Excel .xlsx，可能下載未完成或檔案損壞。\n\n原始錯誤：{error}"
    if isinstance(error, MemoryError):
        return "記憶體不足。請把「平行檔案數」降為 1 或 2，關閉其他大型程式後重跑；已完成日期可由快取續跑。"
    if isinstance(error, OSError):
        error_number = getattr(error, "errno", None)
        windows_error = getattr(error, "winerror", None)
        if error_number == 28 or windows_error == 112:
            return "磁碟空間不足。請清出至少 2 GB，或將輸出位置改到空間較大的磁碟後重跑。"
        return f"檔案系統發生錯誤。請確認來源／輸出磁碟仍連線且有足夠空間。\n\n原始錯誤：{error}"
    if isinstance(error, ValueError):
        return str(error)
    if type(error).__name__ in {"BrokenProcessPool", "BrokenExecutor"}:
        return "平行處理程序異常停止。請將「平行檔案數」降為 1 或 2 後重跑；已完成日期會沿用快取。"
    return (
        "發生未預期錯誤。程式已保留已完成日期的快取；請依錯誤報告中的檔名與訊息處理後重跑。\n\n"
        f"{type(error).__name__}: {error}"
    )


def write_error_report(config: AppConfig, error: BaseException) -> Path | None:
    try:
        config.output_path.parent.mkdir(parents=True, exist_ok=True)
        report = config.output_path.parent / f"AIS月報_錯誤報告_{datetime.now():%Y%m%d_%H%M%S}.txt"
        try:
            discovered, _warnings = discover_source_files(config.input_dir, config.port)
            source_count: int | str = sum(
                1
                for parsed in discovered
                if parsed.year == config.year and parsed.month == config.month
            )
        except (OSError, ValueError):
            source_count = "無法判定"
        contents = [
            APP_TITLE,
            f"版本：{APP_VERSION}",
            f"時間：{datetime.now():%Y-%m-%d %H:%M:%S}",
            f"來源：{config.input_dir}",
            f"港別：{config.port}",
            f"月份：{config.year}-{config.month:02d}",
            f"來源檔數：{source_count}",
            f"新版輸出：{config.output_path}",
            f"原格式輸出：{config.legacy_output_path or '未要求'}",
            "",
            "給使用者的說明：",
            friendly_error_message(error),
            "",
            "技術細節：",
            "".join(traceback.format_exception(type(error), error, error.__traceback__)),
        ]
        report.write_text("\n".join(contents), encoding="utf-8-sig")
        return report
    except OSError:
        return None


def launch_gui() -> None:
    # PyInstaller one-file extracts Tcl/Tk beside the frozen program.  Set the
    # paths explicitly before importing tkinter because some Windows Python
    # distributions fail to apply the bundled runtime hook early enough.
    bundle_root = getattr(sys, "_MEIPASS", None)
    if bundle_root:
        bundle_path = Path(bundle_root)
        os.environ["TCL_LIBRARY"] = str(bundle_path / "_tcl_data")
        os.environ["TK_LIBRARY"] = str(bundle_path / "_tk_data")
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    class Application(tk.Tk):
        def __init__(self) -> None:
            super().__init__()
            self.title(f"{APP_TITLE} v{APP_VERSION}")
            self.geometry("960x760")
            self.minsize(860, 700)
            self.configure(bg="#F4F7F9")
            self.events: queue.Queue[dict] = queue.Queue()
            self.cancel_event = threading.Event()
            self.worker: threading.Thread | None = None
            self.last_output: Path | None = None
            self.last_outputs: list[Path] = []

            self.source_var = tk.StringVar()
            self.output_var = tk.StringVar()
            self.legacy_var = tk.BooleanVar(value=True)
            self.legacy_path_var = tk.StringVar(value="將依新版檔名自動建立")
            self.port_choice_var = tk.StringVar(value="請先選擇每日資料夾")
            self.month_choice_var = tk.StringVar(value="請先選擇每日資料夾")
            self.source_catalog: dict[str, list[tuple[int, int, int]]] = {}
            self.month_options: dict[str, tuple[int, int, int]] = {}
            self.max_distance_var = tk.StringVar(value="500")
            self.tolerance_var = tk.StringVar(value="10")
            self.cluster_var = tk.StringVar(value="3")
            self.top_var = tk.StringVar(value="50")
            self.workers_var = tk.StringVar(value=str(default_worker_count()))
            self.message_types_var = tk.StringVar(value="1,2,3,18,19")
            self.status_var = tk.StringVar(value="請先選擇每日 AIS Excel 所在資料夾。")
            self.progress_text_var = tk.StringVar(value="尚未開始")

            style = ttk.Style(self)
            style.theme_use("vista")
            style.configure("Title.TLabel", background="#17324D", foreground="white", font=("Microsoft JhengHei", 18, "bold"), padding=(18, 14))
            style.configure("Section.TLabelframe", background="#F4F7F9", padding=12)
            style.configure("Section.TLabelframe.Label", background="#F4F7F9", foreground="#17324D", font=("Microsoft JhengHei", 11, "bold"))
            style.configure("Primary.TButton", font=("Microsoft JhengHei", 12, "bold"), padding=(18, 10))
            style.configure("TLabel", background="#F4F7F9", font=("Microsoft JhengHei", 10))
            style.configure("TButton", font=("Microsoft JhengHei", 10))

            ttk.Label(self, text=APP_TITLE, style="Title.TLabel").pack(fill="x")
            container = ttk.Frame(self, padding=16)
            container.pack(fill="both", expand=True)
            container.columnconfigure(0, weight=1)

            files_box = ttk.LabelFrame(container, text="1  檔案位置", style="Section.TLabelframe")
            files_box.grid(row=0, column=0, sticky="ew", pady=(0, 10))
            files_box.columnconfigure(1, weight=1)
            ttk.Label(files_box, text="每日資料夾").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
            self.source_entry = ttk.Entry(files_box, textvariable=self.source_var)
            self.source_entry.grid(row=0, column=1, sticky="ew", pady=5)
            self.source_entry.bind("<Return>", self.refresh_source_from_entry)
            self.source_entry.bind("<FocusOut>", self.refresh_source_from_entry)
            ttk.Button(files_box, text="選擇…", command=self.choose_source).grid(row=0, column=2, padx=(8, 0), pady=5)
            ttk.Label(files_box, text="輸出月報").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=5)
            ttk.Entry(files_box, textvariable=self.output_var).grid(row=1, column=1, sticky="ew", pady=5)
            ttk.Button(files_box, text="另存為…", command=self.choose_output).grid(row=1, column=2, padx=(8, 0), pady=5)
            ttk.Checkbutton(files_box, text="同時產生原格式相容版", variable=self.legacy_var, command=self.update_legacy_path).grid(row=2, column=0, sticky="w", pady=5)
            ttk.Label(files_box, textvariable=self.legacy_path_var, foreground="#5B6573").grid(row=2, column=1, columnspan=2, sticky="w", pady=5)
            self.output_var.trace_add("write", lambda *_args: self.update_legacy_path())

            config_box = ttk.LabelFrame(container, text="2  港別、月份與判斷設定", style="Section.TLabelframe")
            config_box.grid(row=1, column=0, sticky="ew", pady=(0, 10))
            ttk.Label(config_box, text="港別（由檔名判定）").grid(row=0, column=0, sticky="w")
            self.port_combo = ttk.Combobox(
                config_box,
                textvariable=self.port_choice_var,
                state="readonly",
                width=12,
            )
            self.port_combo.grid(row=1, column=0, sticky="w", pady=(3, 8))
            self.port_combo.bind("<<ComboboxSelected>>", self.port_selection_changed)
            ttk.Label(config_box, text="製作月份（由檔名判定）").grid(row=0, column=1, sticky="w", padx=(10, 0))
            self.month_combo = ttk.Combobox(
                config_box,
                textvariable=self.month_choice_var,
                state="readonly",
                width=25,
            )
            self.month_combo.grid(row=1, column=1, sticky="w", padx=(10, 0), pady=(3, 8))
            self.month_combo.bind("<<ComboboxSelected>>", self.month_selection_changed)
            labels = [
                ("距離上限 NM", self.max_distance_var, 8),
                ("群聚差距 %", self.tolerance_var, 7),
                ("至少筆數", self.cluster_var, 5),
                ("每方向候選", self.top_var, 6),
                ("平行檔案數", self.workers_var, 5),
            ]
            for column, (label, variable, width) in enumerate(labels):
                ttk.Label(config_box, text=label).grid(row=2, column=column, sticky="w", padx=(10 if column else 0, 0))
                ttk.Entry(config_box, textvariable=variable, width=width).grid(row=3, column=column, sticky="w", padx=(10 if column else 0, 0), pady=(3, 8))
            ttk.Label(config_box, text="AIS 訊息類型").grid(row=4, column=0, sticky="w")
            ttk.Entry(config_box, textvariable=self.message_types_var, width=24).grid(row=4, column=1, sticky="w", padx=(10, 0))
            ttk.Label(config_box, text="船舶位置預設 1,2,3,18,19；每日完成即保存快取，可中斷後續跑。", foreground="#5B6573").grid(row=4, column=2, columnspan=3, sticky="w", padx=(12, 0))

            run_box = ttk.LabelFrame(container, text="3  一鍵執行", style="Section.TLabelframe")
            run_box.grid(row=2, column=0, sticky="nsew")
            run_box.columnconfigure(0, weight=1)
            container.rowconfigure(2, weight=1)
            button_row = ttk.Frame(run_box)
            button_row.grid(row=0, column=0, sticky="ew")
            self.start_button = ttk.Button(button_row, text="開始全自動製作", style="Primary.TButton", command=self.start)
            self.start_button.pack(side="left")
            self.cancel_button = ttk.Button(button_row, text="取消", command=self.cancel, state="disabled")
            self.cancel_button.pack(side="left", padx=8)
            self.open_button = ttk.Button(button_row, text="開啟輸出資料夾", command=self.open_output, state="disabled")
            self.open_button.pack(side="right")
            ttk.Label(run_box, textvariable=self.status_var, foreground="#17324D").grid(row=1, column=0, sticky="w", pady=(12, 3))
            self.progress = ttk.Progressbar(run_box, mode="determinate", maximum=100)
            self.progress.grid(row=2, column=0, sticky="ew", pady=3)
            ttk.Label(run_box, textvariable=self.progress_text_var, foreground="#5B6573").grid(row=3, column=0, sticky="w", pady=(0, 8))
            self.log = tk.Text(run_box, height=14, wrap="word", font=("Consolas", 9), bg="#FFFFFF", fg="#23313D", relief="solid", borderwidth=1)
            self.log.grid(row=4, column=0, sticky="nsew")
            run_box.rowconfigure(4, weight=1)
            scrollbar = ttk.Scrollbar(run_box, orient="vertical", command=self.log.yview)
            scrollbar.grid(row=4, column=1, sticky="ns")
            self.log.configure(yscrollcommand=scrollbar.set, state="disabled")
            self.after(150, self.poll_events)

        def append_log(self, text: str) -> None:
            self.log.configure(state="normal")
            self.log.insert("end", f"[{datetime.now():%H:%M:%S}] {text}\n")
            self.log.see("end")
            self.log.configure(state="disabled")

        def choose_source(self) -> None:
            selected = filedialog.askdirectory(title="選擇 D&TMOK AIS 每日 Excel 資料夾")
            if not selected:
                return
            source = Path(selected)
            self.source_var.set(str(source))
            self.refresh_detected_sources(source, announce=True, update_output=True)

        def refresh_source_from_entry(self, _event: object | None = None) -> None:
            raw = self.source_var.get().strip()
            if not raw:
                return
            source = Path(raw)
            if source.is_dir():
                self.refresh_detected_sources(source, announce=False, update_output=True)

        def refresh_detected_sources(self, source: Path, *, announce: bool, update_output: bool) -> None:
            catalog = detect_source_catalog(source)
            previous_port = self.port_choice_var.get()
            previous_month = self.month_options.get(self.month_choice_var.get())
            self.source_catalog = catalog
            ports = list(catalog)
            self.port_combo.configure(values=ports)
            if not ports:
                self.port_choice_var.set("找不到港別")
                self.month_options = {}
                self.month_combo.configure(values=[])
                self.month_choice_var.set("找不到固定格式的每日檔")
                self.status_var.set("找不到符合 D&TMOK <PORT>_YYYYMMDD_*.xlsx 格式的檔案。")
                if announce:
                    self.append_log(self.status_var.get())
                return

            selected_port = previous_port if previous_port in catalog else ports[0]
            self.port_choice_var.set(selected_port)
            self._refresh_month_options(
                selected_port,
                previous_month=previous_month if selected_port == previous_port else None,
                update_output=update_output,
            )
            selected = self.month_options.get(self.month_choice_var.get())
            if selected is None:
                return
            year, month, count = selected
            if len(ports) == 1:
                port_message = f"已自動辨識唯一港別 {selected_port}。"
            else:
                port_message = f"偵測到 {len(ports)} 個港別（{'、'.join(ports)}）；目前選擇 {selected_port}。"
            month_message = (
                f"{year} 年 {month} 月，共 {count} 個每日檔。"
                if len(catalog[selected_port]) == 1
                else f"已選最新的 {year} 年 {month} 月，共 {count} 個每日檔；可由清單改選。"
            )
            message = f"{port_message} {month_message}"
            self.status_var.set(message)
            if announce:
                self.append_log(message)

        def _refresh_month_options(
            self,
            port: str,
            *,
            previous_month: tuple[int, int, int] | None = None,
            update_output: bool,
        ) -> None:
            months = self.source_catalog.get(port, [])
            self.month_options = {
                month_option_label(year, month, count): (year, month, count)
                for year, month, count in months
            }
            self.month_combo.configure(values=list(self.month_options))
            if not months:
                self.month_choice_var.set("找不到固定格式的每日檔")
                return
            selected = previous_month if previous_month in months else max(months, key=lambda item: (item[0], item[1]))
            self.month_choice_var.set(month_option_label(*selected))
            year, month, _count = selected
            if update_output:
                default_name = default_output_filename(port, year, month)
                self.output_var.set(str(default_output_directory() / default_name))

        def port_selection_changed(self, _event: object | None = None) -> None:
            port = self.port_choice_var.get()
            if port not in self.source_catalog:
                return
            self._refresh_month_options(port, update_output=True)
            selected = self.month_options.get(self.month_choice_var.get())
            if selected is None:
                return
            year, month, count = selected
            message = f"已選定港別 {port}；月份清單已更新為該港別資料，預設 {year} 年 {month} 月（{count} 個每日檔）。"
            self.status_var.set(message)
            self.append_log(message)

        def month_selection_changed(self, _event: object | None = None) -> None:
            selected = self.month_options.get(self.month_choice_var.get())
            if selected is None:
                return
            year, month, count = selected
            port = self.port_choice_var.get()
            default_name = default_output_filename(port, year, month)
            self.output_var.set(str(default_output_directory() / default_name))
            self.status_var.set(f"已由檔名選定 {port} / {year} 年 {month} 月，共 {count} 個每日檔。")
            self.append_log(self.status_var.get())

        def choose_output(self) -> None:
            initial = Path(self.output_var.get()) if self.output_var.get() else Path.cwd() / "AIS_32方位數值.xlsx"
            selected = filedialog.asksaveasfilename(
                title="儲存 AIS 月報",
                initialdir=str(initial.parent),
                initialfile=initial.name,
                defaultextension=".xlsx",
                filetypes=[("Excel 活頁簿", "*.xlsx")],
            )
            if selected:
                self.output_var.set(selected)

        def update_legacy_path(self) -> None:
            raw = self.output_var.get().strip()
            if not self.legacy_var.get():
                self.legacy_path_var.set("不產生原格式相容版")
            elif raw:
                self.legacy_path_var.set(str(derive_legacy_output_path(Path(raw))))
            else:
                self.legacy_path_var.set("將依新版檔名自動建立")

        def build_config(self) -> AppConfig:
            source_text = self.source_var.get().strip()
            output_text = self.output_var.get().strip()
            if not source_text:
                raise ValueError("請選擇每日資料夾。")
            if not output_text:
                raise ValueError("請指定輸出月報。")
            source = Path(source_text)
            output = Path(output_text)
            detected = detect_source_catalog(source)
            if not detected:
                raise ValueError("找不到符合 D&TMOK <PORT>_YYYYMMDD_*.xlsx 格式的檔案。")
            port = self.port_choice_var.get()
            if port not in detected:
                self.refresh_detected_sources(source, announce=True, update_output=False)
                raise ValueError("來源資料夾內容已變更，原選擇的港別已不存在；請重新確認港別與月份。")
            selected = self.month_options.get(self.month_choice_var.get())
            if selected not in detected[port]:
                self.refresh_detected_sources(source, announce=True, update_output=False)
                raise ValueError("來源資料夾內容已變更，原選擇的月份已不存在；請重新確認港別與月份。")
            year, month, _count = selected
            legacy_output = derive_legacy_output_path(output) if self.legacy_var.get() else None
            return AppConfig(
                input_dir=source,
                output_path=output,
                port=port,
                year=year,
                month=month,
                legacy_output_path=legacy_output,
                max_distance=float(self.max_distance_var.get()),
                tolerance=float(self.tolerance_var.get()) / 100.0,
                cluster_size=int(self.cluster_var.get()),
                top_candidates=int(self.top_var.get()),
                message_types=parse_message_types(self.message_types_var.get()),
                workers=int(self.workers_var.get()),
                overwrite=True,
            )

        def start(self) -> None:
            try:
                config = self.build_config()
                existing = [path for path in (config.output_path, config.legacy_output_path) if path is not None and path.exists()]
                if existing:
                    listing = "\n".join(str(path) for path in existing)
                    if not messagebox.askyesno("覆寫確認", f"下列檔案已存在：\n{listing}\n\n要覆寫嗎？"):
                        return
                config.validate()
            except Exception as error:
                messagebox.showerror("設定錯誤", str(error))
                return
            self.cancel_event.clear()
            self.last_output = None
            self.last_outputs = []
            self.start_button.configure(state="disabled")
            self.cancel_button.configure(state="normal")
            self.open_button.configure(state="disabled")
            self.progress["value"] = 0
            self.status_var.set("開始處理…")
            self.progress_text_var.set("正在準備來源檔")
            selected = self.month_options.get(self.month_choice_var.get())
            source_count = selected[2] if selected is not None else 0
            self.append_log(f"Port: {config.port}")
            self.append_log(f"Period: {config.year}-{config.month:02d}")
            self.append_log(f"Source files: {source_count}")
            self.append_log(f"新版：{config.output_path}")
            if config.legacy_output_path is not None:
                self.append_log(f"原格式版：{config.legacy_output_path}")

            def callback(payload: dict) -> None:
                self.events.put(payload)

            def work() -> None:
                try:
                    output = run_pipeline(config, callback=callback, cancel_event=self.cancel_event)
                    outputs = [str(output)]
                    if config.legacy_output_path is not None:
                        outputs.append(str(config.legacy_output_path))
                    self.events.put({"kind": "worker_success", "output": str(output), "outputs": outputs})
                except CancelledError as error:
                    self.events.put({"kind": "worker_cancelled", "error": str(error)})
                except Exception as error:
                    report = write_error_report(config, error)
                    self.events.put(
                        {
                            "kind": "worker_error",
                            "error": friendly_error_message(error),
                            "technical": f"{type(error).__name__}: {error}",
                            "report": str(report) if report else None,
                        }
                    )

            self.worker = threading.Thread(target=work, daemon=True)
            self.worker.start()

        def cancel(self) -> None:
            self.cancel_event.set()
            self.cancel_button.configure(state="disabled")
            self.status_var.set("正在安全停止；目前來源檔讀完或每 10,000 列會檢查一次…")
            self.append_log("已要求取消。")

        def poll_events(self) -> None:
            try:
                while True:
                    payload = self.events.get_nowait()
                    kind = payload.get("kind")
                    if kind == "file_start":
                        position, total = int(payload["position"]), int(payload["total"])
                        self.progress["value"] = (position - 1) / max(total, 1) * 85
                        self.status_var.set(f"處理第 {position}/{total} 檔：{payload['file']}")
                        self.progress_text_var.set("正在逐列讀取並保留各方位候選值")
                        self.append_log(self.status_var.get())
                    elif kind == "preflight":
                        required = int(payload["required_bytes"]) / 1024**3
                        free_value = payload.get("free_bytes")
                        free_text = f"，可用 {int(free_value) / 1024**3:.1f} GB" if free_value else ""
                        self.append_log(f"輸出空間預估至少 {required:.1f} GB{free_text}")
                    elif kind == "parallel_start":
                        self.status_var.set(f"使用 {payload['workers']} 個處理程序平行讀取 {payload['total']} 個來源檔…")
                        self.append_log(self.status_var.get())
                    elif kind == "cache_hit":
                        position, total = int(payload["position"]), int(payload["total"])
                        self.progress["value"] = position / max(total, 1) * 85
                        self.progress_text_var.set(f"沿用已完成快取：{payload['file']}")
                        self.append_log(f"快取命中：{payload['file']}")
                    elif kind == "rows":
                        rows, estimated = int(payload["rows"]), int(payload["estimated_rows"])
                        self.progress_text_var.set(f"{payload['file']}：已掃描 {rows:,} / 約 {estimated:,} 列")
                    elif kind == "file_done":
                        position, total = int(payload["position"]), int(payload["total"])
                        self.progress["value"] = position / max(total, 1) * 85
                        self.append_log(f"完成 {payload['file']}，耗時 {format_seconds(float(payload['seconds']))}")
                    elif kind == "write_start":
                        self.progress["value"] = 88
                        self.status_var.set("資料讀取完成，正在製作 Excel 月報與圖表…")
                        self.progress_text_var.set(payload["file"])
                        self.append_log(self.status_var.get())
                    elif kind == "sheet_written":
                        position, total = int(payload["position"]), int(payload["total"])
                        self.progress["value"] = 88 + position / max(total, 1) * 10
                        self.progress_text_var.set(f"已建立 {payload['sheet']}（{position}/{total}）")
                    elif kind == "write_done":
                        self.progress["value"] = 92 if self.legacy_var.get() else 99
                        self.append_log(f"Excel 已寫入：{payload['file']}")
                    elif kind == "legacy_write_start":
                        self.progress["value"] = 92
                        self.status_var.set("新版完成，正在製作原格式完整資料版…")
                        self.progress_text_var.set(payload["file"])
                        self.append_log(self.status_var.get())
                    elif kind == "legacy_sheet_written":
                        position, total = int(payload["position"]), int(payload["total"])
                        self.progress["value"] = 92 + position / max(total, 1) * 7
                        self.progress_text_var.set(
                            f"原格式：{payload['sheet']}，寫入 {int(payload['rows']):,} 列（{position}/{total}）"
                        )
                    elif kind == "legacy_write_done":
                        self.progress["value"] = 99
                        self.append_log(f"原格式 Excel 已寫入：{payload['file']}")
                    elif kind == "complete":
                        self.progress_text_var.set(f"總耗時：{format_seconds(float(payload['seconds']))}")
                    elif kind == "worker_success":
                        self.last_output = Path(payload["output"])
                        self.last_outputs = [Path(value) for value in payload.get("outputs", [payload["output"]])]
                        self.progress["value"] = 100
                        self.status_var.set("完成。新版分析版與原格式相容版均已產生。" if len(self.last_outputs) == 2 else "完成。新版分析月報已產生。")
                        for completed_output in self.last_outputs:
                            self.append_log(f"完成：{completed_output}")
                        self.start_button.configure(state="normal")
                        self.cancel_button.configure(state="disabled")
                        self.open_button.configure(state="normal")
                        listing = "\n".join(str(path) for path in self.last_outputs)
                        messagebox.showinfo("製作完成", f"AIS 月報已完成：\n{listing}\n\n請先查看新版的「待複核」工作表。")
                    elif kind == "worker_cancelled":
                        self.status_var.set("已取消；未覆寫正式輸出檔。")
                        self.append_log(payload["error"])
                        self.start_button.configure(state="normal")
                        self.cancel_button.configure(state="disabled")
                    elif kind == "worker_error":
                        self.status_var.set("製作失敗；請看下方錯誤。")
                        self.append_log(payload.get("technical", payload["error"]))
                        if payload.get("report"):
                            self.append_log(f"錯誤報告：{payload['report']}")
                        self.start_button.configure(state="normal")
                        self.cancel_button.configure(state="disabled")
                        report_note = f"\n\n完整錯誤報告：\n{payload['report']}" if payload.get("report") else ""
                        messagebox.showerror("製作失敗", payload["error"] + report_note)
            except queue.Empty:
                pass
            self.after(150, self.poll_events)

        def open_output(self) -> None:
            if self.last_output:
                subprocess.Popen(["explorer", "/select,", str(self.last_output)])

    Application().mainloop()


def cli_progress(payload: dict) -> None:
    if sys.stdout is None:
        return
    kind = payload.get("kind")
    if kind == "job":
        print(f"Port: {payload['port']}", flush=True)
        print(f"Period: {payload['year']}-{int(payload['month']):02d}", flush=True)
        print(f"Source files: {payload['source_count']}", flush=True)
    elif kind == "file_start":
        print(f"[{payload['position']}/{payload['total']}] {payload['file']}", flush=True)
    elif kind == "rows":
        print(f"  scanned {int(payload['rows']):,} rows", flush=True)
    elif kind == "file_done":
        print(f"  done in {format_seconds(float(payload['seconds']))}", flush=True)
    elif kind == "write_start":
        print("Writing workbook and charts…", flush=True)
    elif kind == "legacy_write_start":
        print("Writing legacy-compatible full workbook…", flush=True)
    elif kind == "legacy_sheet_written":
        print(
            f"  legacy {payload['position']}/{payload['total']} {payload['sheet']}: {int(payload['rows']):,} rows",
            flush=True,
        )
    elif kind == "complete":
        print(f"Complete in {format_seconds(float(payload['seconds']))}: {payload['file']}", flush=True)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=APP_TITLE)
    parser.add_argument("--input", type=Path, help="每日 D&TMOK <PORT> Excel 資料夾")
    parser.add_argument("--output", type=Path, help="輸出 .xlsx")
    parser.add_argument("--legacy-output", type=Path, help="同時輸出的原格式相容版 .xlsx")
    parser.add_argument("--port", help="可省略；來源含多港別時必須指定，例如 HWLN")
    parser.add_argument("--year", type=int, help="可省略；預設從固定格式檔名自動判定")
    parser.add_argument("--month", type=int, help="可省略；預設從固定格式檔名自動判定")
    parser.add_argument("--max-distance", type=float, default=500.0)
    parser.add_argument("--tolerance", type=float, default=10.0, help="百分比，例如 10")
    parser.add_argument("--cluster-size", type=int, default=3)
    parser.add_argument("--top-candidates", type=int, default=50)
    parser.add_argument("--message-types", default="1,2,3,18,19")
    parser.add_argument(
        "--workers",
        type=int,
        default=default_worker_count(),
        help="同時處理的每日檔案數；預設依電腦自動使用 1–3",
    )
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--max-files", type=int, help="測試用：只處理前 N 檔")
    arguments = parser.parse_args(argv)

    if all(value is None for value in (arguments.input, arguments.output, arguments.legacy_output, arguments.port, arguments.year, arguments.month)):
        launch_gui()
        return 0
    if arguments.input is None or arguments.output is None:
        parser.error("命令列模式必須同時指定 --input 與 --output；年份、月份可由檔名自動判定")
    try:
        port, year, month, count = resolve_source_period(
            arguments.input,
            arguments.port,
            arguments.year,
            arguments.month,
        )
    except ValueError as error:
        parser.error(str(error))
    print(f"Filename job detected: {port} / {year}-{month:02d} ({count} daily files)", flush=True)
    config = AppConfig(
        input_dir=arguments.input,
        output_path=arguments.output,
        port=port,
        year=year,
        month=month,
        legacy_output_path=arguments.legacy_output,
        max_distance=arguments.max_distance,
        tolerance=arguments.tolerance / 100.0,
        cluster_size=arguments.cluster_size,
        top_candidates=arguments.top_candidates,
        message_types=parse_message_types(arguments.message_types),
        workers=arguments.workers,
        overwrite=arguments.overwrite,
        max_files=arguments.max_files,
    )
    run_pipeline(config, callback=cli_progress)
    return 0


if __name__ == "__main__":
    multiprocessing.freeze_support()
    raise SystemExit(main())
