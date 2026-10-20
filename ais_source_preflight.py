"""Read-only, full-row source safety gate. Never selects AIS distances.

GPS fields may be inspected as forensic timestamps, but are never accepted as
production time aliases. No input is rewritten, reassigned, unioned or excluded.
"""
from __future__ import annotations

import calendar
import concurrent.futures
import hashlib
import json
import math
import posixpath
import struct
import time
import xml.etree.ElementTree as ET
import xml.sax
from array import array
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path
from xml.sax.handler import ContentHandler


PREFLIGHT_SCHEMA = "AIS_SOURCE_PREFLIGHT_1"
GPS_TIME = {key: "gps_" + key for key in ("year", "month", "day", "hour", "minute", "second")}
CORE_STRUCT = struct.Struct("<IIddq?")


class PreflightCancelled(RuntimeError):
    pass


def sha256(path):
    with Path(path).open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest().upper()


def normalized_header(value):
    return " ".join(str(value or "").strip().casefold().split())


def _number(value):
    try:
        number = float(value)
    except (ValueError, TypeError, OverflowError):
        return None
    return number if math.isfinite(number) else None


def _integer(value):
    number = _number(value)
    return int(number) if number is not None and number.is_integer() else None


def _column(reference):
    result = 0
    for character in reference:
        if not character.isalpha():
            break
        result = result * 26 + ord(character.upper()) - 64
    return result - 1


class _HeaderDone(Exception):
    pass


class _Rows(ContentHandler):
    def __init__(self, shared, callback, wanted=None, header_only=False):
        super().__init__()
        self.shared, self.callback, self.wanted, self.header_only = shared, callback, wanted, header_only
        self.row = None
        self.row_number = 0
        self.column = -1
        self.kind = None
        self.capture = False
        self.text_parts = []
        self.declared_last_row = None

    def startElement(self, name, attributes):
        name = name.split(":")[-1]
        if name == "dimension":
            last = attributes.get("ref","").split(":")[-1]
            digits = "".join(c for c in last if c.isdigit())
            self.declared_last_row = int(digits) if digits else None
        elif name == "row":
            self.row_number = int(attributes.get("r", self.row_number + 1))
            self.row = {}
        elif name == "c" and self.row is not None:
            self.column = _column(attributes.get("r", ""))
            self.kind = attributes.get("t", "n")
            self.text_parts = []
        elif name in ("v", "t") and self.row is not None:
            self.capture = self.wanted is None or self.row_number == 1 or self.column in self.wanted

    def characters(self, content):
        if self.capture:
            self.text_parts.append(content)

    def endElement(self, name):
        name = name.split(":")[-1]
        if name in ("v", "t"):
            self.capture = False
        elif name == "c" and self.row is not None:
            if self.wanted is None or self.row_number == 1 or self.column in self.wanted:
                text = "".join(self.text_parts)
                if self.kind == "s":
                    value = self.shared[int(text)] if text else None
                elif self.kind == "b":
                    value = text == "1"
                elif self.kind == "n":
                    # Preserve the production reader's XML numeric identity:
                    # <v>1</v> is int 1, while <v>1.0</v> is float 1.0.
                    # Converting both to float changes str(channel) and can
                    # miss duplicate cores against a text channel "1".
                    value = (float(text) if any(c in text for c in ".Ee") else int(text)) if text else None
                else:
                    value = text or None
                self.row[self.column] = value
        elif name == "row" and self.row is not None:
            self.callback(self.row_number, self.row)
            self.row = None
            if self.header_only:
                raise _HeaderDone()


def _parse(archive, target, shared, callback, wanted=None, header_only=False):
    parser = xml.sax.make_parser()
    parser.setFeature(xml.sax.handler.feature_external_ges, False)
    handler = _Rows(shared, callback, wanted, header_only)
    parser.setContentHandler(handler)
    try:
        with archive.open(target) as stream:
            parser.parse(stream)
    except _HeaderDone:
        pass
    return handler.declared_last_row


def _workbook_parts(archive):
    ns = {"s":"http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    rel = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    relationships = {r.attrib["Id"]:r.attrib["Target"]
                     for r in ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))}
    sheets = []
    for element in ET.fromstring(archive.read("xl/workbook.xml")).findall("s:sheets/s:sheet", ns):
        target = relationships[element.attrib[rel]]
        target = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/"+target)
        if not target.startswith("xl/") or ".." in target.split("/"):
            raise ValueError("Unsafe workbook relationship")
        sheets.append((element.attrib["name"], target))
    shared = []
    if "xl/sharedStrings.xml" in archive.namelist():
        with archive.open("xl/sharedStrings.xml") as stream:
            for _event, element in ET.iterparse(stream, events=("end",)):
                if element.tag.endswith("}si"):
                    shared.append("".join(t.text or "" for t in element.iter() if t.tag.endswith("}t")))
                    element.clear()
    return sorted(sheets, key=lambda item:item[0].strip().casefold() != "ais"), shared


def inspect_headers(header, aliases):
    matches = {key:[column for column,value in header.items() if normalized_header(value) in names]
               for key,names in aliases.items()}
    gps = {key:[column for column,value in header.items() if normalized_header(value) == label]
           for key,label in GPS_TIME.items()}
    any_gps = any(gps.values())
    all_gps = all(len(v)==1 for v in gps.values())
    standard_time = any(matches[key] for key in GPS_TIME)
    if any_gps:
        profile = "GPS_SCHEMA_DETECTED" if all_gps and not standard_time else "AMBIGUOUS_TIME_SCHEMA"
    elif all(len(v)==1 for v in matches.values()):
        profile = "STANDARD_AIS"
    else:
        profile = "UNKNOWN_OR_INCOMPLETE_SCHEMA"
    indexes = {key:values[0] for key,values in matches.items() if len(values)==1}
    if profile == "GPS_SCHEMA_DETECTED":
        indexes.update({key:values[0] for key,values in gps.items()})
    missing = [key for key in aliases if key not in indexes]
    duplicate = {key:columns for key,columns in matches.items() if len(columns)>1}
    fingerprint = hashlib.sha256(json.dumps(sorted(header.items()),ensure_ascii=False).encode()).hexdigest().upper()
    return {"profile":profile,"indexes":indexes,"missing_fields":missing,"duplicate_fields":duplicate,"header_fingerprint_sha256":fingerprint,
            "headers":{str(k+1):v for k,v in sorted(header.items())},
            "time_basis":"GPS_FIELDS_FORENSIC_ONLY_NOT_AUTHORIZED" if any_gps else "STANDARD_FIELDS_AS_RECORDED"}


def _scan_file(source, aliases, keep_records, cancel_event=None):
    import zipfile
    path = Path(source["path"])
    result = {**source,"state":"SCANNING","issues":[],"rows_scanned":0,"timestamp_dates":{},
              "mixed_date_rows":0,"invalid_time_rows":0,"normalized_core_rows":0,"invalid_numeric_rows":0,
              "not_east_rows":0,"min_second":None,"max_second":None,"examples":[],"worksheets":[],"row_coverage":"UNVERIFIED"}
    occupied = bytearray(86400)
    counts_by_second = array("I", [0])*86400
    exact = Counter()
    events = Counter()
    expected_day = date.fromisoformat(source["filename_date"])
    previous_row = 1
    row_gaps = 0
    dates = Counter()
    try:
        before = path.stat()
        result["sha256"] = sha256(path)
        with zipfile.ZipFile(path) as archive:
            sheets, shared = _workbook_parts(archive)
            candidates = []
            for title,target in sheets:
                header = {}
                def read_header(number, values):
                    if number == 1:
                        header.update(values)
                _parse(archive, target, shared, read_header, header_only=True)
                info = inspect_headers(header, aliases)
                result["worksheets"].append({"name":title,**info})
                if not info["missing_fields"] and not info["duplicate_fields"] and info["profile"] in ("STANDARD_AIS","GPS_SCHEMA_DETECTED"):
                    candidates.append((title,target,info))
            if len(candidates)!=1:
                result["issues"].append("AMBIGUOUS_DATA_WORKSHEETS" if len(candidates)>1 else "SCHEMA_MISMATCH")
                if any(s["profile"]=="GPS_SCHEMA_DETECTED" for s in result["worksheets"]):
                    result["issues"].append("GPS_SCHEMA_DETECTED")
                result["schema_profile"] = "UNVERIFIED"
            else:
                title,target,info = candidates[0]
                result["selected_sheet"],result["schema_profile"] = title,info["profile"]
                result["field_positions_1based"] = {k:v+1 for k,v in info["indexes"].items()}
                result["time_basis"] = info["time_basis"]
                result["header_fingerprint_sha256"] = info["header_fingerprint_sha256"]
                if info["profile"]=="GPS_SCHEMA_DETECTED":
                    result["issues"].append("GPS_SCHEMA_DETECTED")
                indexes = info["indexes"]
                def row(number, values):
                    nonlocal previous_row, row_gaps
                    if number==1:
                        return
                    if cancel_event is not None and cancel_event.is_set():
                        raise PreflightCancelled("Source preflight cancelled")
                    row_gaps += max(0,number-previous_row-1)
                    previous_row = number
                    result["rows_scanned"] += 1
                    parts = [_integer(values.get(indexes[key])) for key in GPS_TIME]
                    try:
                        if any(v is None for v in parts):
                            raise ValueError("Non-integer timestamp field")
                        timestamp = datetime(*parts)
                    except (ValueError,TypeError,OverflowError) as error:
                        result["invalid_time_rows"] += 1
                        if len(result["examples"])<12:
                            result["examples"].append({"source_row":number,"kind":"INVALID_TIMESTAMP","reason":str(error)})
                        return
                    dates[timestamp.date().isoformat()] += 1
                    second = timestamp.hour*3600+timestamp.minute*60+timestamp.second
                    occupied[second] = 1
                    result["min_second"] = second if result["min_second"] is None else min(second,result["min_second"])
                    result["max_second"] = second if result["max_second"] is None else max(second,result["max_second"])
                    if timestamp.date()!=expected_day:
                        result["mixed_date_rows"] += 1
                        if len(result["examples"])<12:
                            result["examples"].append({"source_row":number,"kind":"MIXED_DATE","actual_date":timestamp.date().isoformat()})
                        return
                    msg = _integer(values.get(indexes["msg_type"]))
                    if msg is None or not 0<=msg<=0xFFFFFFFF:
                        result["invalid_numeric_rows"] += 1
                        return
                    if str(values.get(indexes["longitude_desc"])).strip().casefold()!="east":
                        result["not_east_rows"] += 1
                        return
                    distance = _number(values.get(indexes["distance"]))
                    bearing = _number(values.get(indexes["bearing"]))
                    if distance is None or distance<0 or bearing is None:
                        result["invalid_numeric_rows"] += 1
                        return
                    mmsi = _integer(values.get(indexes["mmsi"]))
                    raw_channel = values.get(indexes["channel"])
                    channel = ("" if raw_channel is None else str(raw_channel).strip()).encode("utf-8")
                    if len(channel)>16 or (mmsi is not None and not -(2**63)<=mmsi<2**63):
                        if "SPOOL_IDENTITY_OVERFLOW" not in result["issues"]:
                            result["issues"].append("SPOOL_IDENTITY_OVERFLOW")
                        return
                    result["normalized_core_rows"] += 1
                    counts_by_second[second] += 1
                    if keep_records:
                        # Exact normalized core multiset, including multiplicity.
                        # No source row is picked or discarded by this evidence.
                        key = CORE_STRUCT.pack(second,msg,bearing%360.0,distance,mmsi or 0,mmsi is None)+channel
                        exact[key] += 1
                        if mmsi is not None:
                            events[(second,msg,mmsi)] += 1
                declared_last = _parse(archive, target, shared, row, set(indexes.values()))
                result["declared_last_row"] = declared_last
                result["physical_last_row"] = previous_row
                if declared_last is not None and declared_last!=previous_row:
                    result["issues"].append("WORKSHEET_DIMENSION_MISMATCH")
                result["row_coverage"] = "COMPLETE_RECOGNIZED_WORKSHEET"
        after = path.stat()
        if (before.st_size,before.st_mtime_ns)!=(after.st_size,after.st_mtime_ns) or sha256(path)!=result["sha256"]:
            result["issues"].append("SOURCE_CHANGED_DURING_PREFLIGHT")
        result["bytes"],result["mtime_ns"] = after.st_size,after.st_mtime_ns
        result["timestamp_dates"] = dict(sorted(dates.items()))
        result["row_gaps"] = row_gaps
        if row_gaps:
            result["issues"].append("EMPTY_OR_MISSING_DATA_ROWS")
        if result["mixed_date_rows"]:
            result["issues"].append("MIXED_DATE_ROWS")
        if result["invalid_time_rows"]:
            result["issues"].append("INVALID_TIMESTAMP_ROWS")
        if not result["rows_scanned"] and not result["issues"]:
            result["issues"].append("NO_DATA_ROWS")
        result["occupied_seconds"] = sum(occupied)
        result["partial_day_indicators"] = []
        if result["min_second"] is not None:
            if result["min_second"]>0:
                result["partial_day_indicators"].append("FIRST_OBSERVED_TIMESTAMP_AFTER_00_00_00")
            if result["max_second"]<86399:
                result["partial_day_indicators"].append("LAST_OBSERVED_TIMESTAMP_BEFORE_23_59_59")
        result["coverage_boundary"] = "Observed timestamp envelope only; not proof of continuous capture or complete 24-hour coverage."
        result["state"] = "BLOCKED" if result["issues"] else "READABLE"
    except PreflightCancelled:
        raise
    except Exception as error:
        result["state"] = "BLOCKED"
        result["issues"].append("WORKBOOK_READ_FAILURE")
        result["read_error"] = f"{type(error).__name__}: {error}"
    return result,occupied,counts_by_second,exact,events


def _scan_day(arguments, cancel_event=None):
    sources,aliases,limit = arguments
    started = time.perf_counter()
    scans = [_scan_file(s,aliases,len(sources)>1,cancel_event) for s in sources]
    pairs = []
    for i,(left,lo,lc,lx,le) in enumerate(scans):
        for right,ro,rc,rx,re in scans[i+1:]:
            overlap_seconds = sum(bool(a and b) for a,b in zip(lo,ro))
            overlap_records = sum(a+b for a,b,x,y in zip(lc,rc,lo,ro) if x and y)
            common = sum((lx & rx).values())
            event_common = sum((le & re).values())
            labels = []
            blocked = False
            if left.get("sha256") and left.get("sha256")==right.get("sha256"):
                labels.append("EXACT_FILE_DUPLICATE")
                blocked = True
            if lx and rx and common==min(sum(lx.values()),sum(rx.values())):
                labels.append("CORE_ROWS_EXACT_DUPLICATE" if lx==rx else "SUBSET_CANDIDATE")
                blocked = True
            elif common:
                labels.append("EXACT_NORMALIZED_CORE_OVERLAP")
                blocked = True
            if overlap_records>limit:
                labels.append("PRODUCTION_OVERLAP_AUDIT_LIMIT")
                blocked = True
            if overlap_seconds:
                labels.append("OCCUPIED_SECOND_OVERLAP")
                smaller_occupied = min(sum(lo),sum(ro))
                if smaller_occupied and overlap_seconds/smaller_occupied>=0.9:
                    labels.append("HIGH_OCCUPIED_SECOND_OVERLAP_INDICATOR")
            if event_common>common:
                labels.append("POSSIBLE_DUPLICATE_EVENTS")
            if not labels:
                labels.append("COMPLEMENTARY_FRAGMENTS")
            pair = {"left":left["path"],"right":right["path"],"left_sha256":left.get("sha256"),
                    "right_sha256":right.get("sha256"),"classifications":labels,"blocked":blocked,
                    "occupied_second_overlap":overlap_seconds,"overlap_normalized_rows":overlap_records,
                    "exact_normalized_core_rows_shared":common,"possible_duplicate_events":max(0,event_common-common),
                    "occupied_second_overlap_fraction_of_smaller":overlap_seconds/min(sum(lo),sum(ro)) if min(sum(lo),sum(ro)) else None,
                    "core_overlap_fraction_of_smaller":common/min(sum(lx.values()),sum(rx.values())) if min(sum(lx.values()),sum(rx.values())) else None,
                    "left_normalized_rows":sum(lx.values()),"right_normalized_rows":sum(rx.values()),
                    "subset_evidence_scope":"Complete normalized required-field multiset including multiplicity, not all unused worksheet fields.",
                    "automatic_exclusion":False}
            if "SUBSET_CANDIDATE" in labels:
                pair["subset_path"] = left["path"] if sum(lx.values())<sum(rx.values()) else right["path"]
                pair["superset_path"] = right["path"] if pair["subset_path"]==left["path"] else left["path"]
            pairs.append(pair)
    return {"files":[scan[0] for scan in scans],"fragment_pairs":pairs,"elapsed_seconds":time.perf_counter()-started}


def preflight_sources(sources, aliases, *, workers=1, overlap_limit=250000, discovery_issues=(), ignored_files=(),
                      physical_xlsx_count=None, callback=None, cancel_event=None):
    """Inspect every supplied file; one unsafe file blocks its entire port-month."""
    started = time.perf_counter()
    if cancel_event is not None and cancel_event.is_set():
        raise PreflightCancelled("Source preflight cancelled")
    groups = defaultdict(list)
    for source in sources:
        groups[(source["port"],source["filename_date"])].append(source)
    arguments = [(v,aliases,overlap_limit) for _k,v in sorted(groups.items())]
    results = []
    def accept(result):
        results.append(result)
        if callback:
            callback({"kind":"source_preflight_progress","completed_days":len(results),"total_days":len(arguments),
                      "files":[f["path"] for f in result["files"]]})
    if workers>1 and len(arguments)>1:
        with concurrent.futures.ProcessPoolExecutor(max_workers=workers) as executor:
            pending = [executor.submit(_scan_day,a) for a in arguments]
            for future in concurrent.futures.as_completed(pending):
                if cancel_event is not None and cancel_event.is_set():
                    for f in pending:
                        f.cancel()
                    raise PreflightCancelled("Source preflight cancelled")
                accept(future.result())
    else:
        for argument in arguments:
            if cancel_event is not None and cancel_event.is_set():
                raise PreflightCancelled("Source preflight cancelled")
            accept(_scan_day(argument,cancel_event))
    files = sorted([f for r in results for f in r["files"]],key=lambda f:f["path"].casefold())
    pairs = sorted([p for r in results for p in r["fragment_pairs"]],key=lambda p:(p["left"],p["right"]))
    for record in files:
        record["blocking_group_issues"] = sorted({label for pair in pairs if pair["blocked"] and record["path"] in (pair["left"],pair["right"])
                                                  for label in pair["classifications"]})
        record["processing_status"] = "BLOCKED" if record["state"]=="BLOCKED" or record["blocking_group_issues"] else "ELIGIBLE"
    months = defaultdict(list)
    for f in files:
        months[(f["port"],f["filename_date"][:7])].append(f)
    month_reports = []
    for (port,month),members in sorted(months.items()):
        year,number = map(int,month.split("-"))
        calendar_days = [date(year,number,d).isoformat() for d in range(1,calendar.monthrange(year,number)[1]+1)]
        source_days = sorted({f["filename_date"] for f in members})
        names = {f["path"] for f in members}
        month_pairs = [p for p in pairs if p["left"] in names]
        bad_pairs = [p for p in month_pairs if p["blocked"]]
        blocked = [f["path"] for f in members if f["processing_status"]=="BLOCKED"]
        missing = sorted(set(calendar_days)-set(source_days))
        month_reports.append({"port":port,"month":month,"physical_xlsx_count":len(members),
                              "logical_dates":source_days,"source_days":len(source_days),"calendar_days":len(calendar_days),
                              "missing_days":missing,"incomplete_month":bool(missing),
                              "multiple_fragment_dates":sorted(d for d in source_days if sum(f["filename_date"]==d for f in members)>1),
                              "observed_timestamp_dates":sorted({d for f in members for d in f["timestamp_dates"]}),
                              "mixed_date_rows":sum(f["mixed_date_rows"] for f in members),
                              "partial_day_files":[f["path"] for f in members if f.get("partial_day_indicators")],
                              "schema_profiles":dict(Counter(f.get("schema_profile","UNVERIFIED") for f in members)),
                              "header_variants":dict(Counter(f.get("header_fingerprint_sha256","UNVERIFIED") for f in members)),
                              "blocked_files":blocked,"blocking_fragment_pairs":bad_pairs,
                              "status":"BLOCKED" if blocked or bad_pairs or discovery_issues else "SAFE_WITH_WARNINGS" if missing or any(f.get("partial_day_indicators") for f in members) or any(p["classifications"]!=["COMPLEMENTARY_FRAGMENTS"] for p in month_pairs) else "SAFE"})
    blocked_months = [f"{m['port']}_{m['month'].replace('-','')}" for m in month_reports if m["status"]=="BLOCKED"]
    blocked = bool(discovery_issues or blocked_months or not files)
    return {"schema":PREFLIGHT_SCHEMA,"created":datetime.now().astimezone().isoformat(),
            "status":"BLOCKED" if blocked else "SAFE_WITH_WARNINGS" if any(m["status"]=="SAFE_WITH_WARNINGS" for m in month_reports) else "SAFE",
            "physical_xlsx_count":physical_xlsx_count if physical_xlsx_count is not None else len(files),
            "applicable_source_files":len(files),"logical_dates":len(groups),"source_days":len(groups),
            "calendar_days":sum(m["calendar_days"] for m in month_reports),
            "rows_scanned":sum(f["rows_scanned"] for f in files),
            "observed_timestamp_dates":[{"port":port,"date":day} for port,day in sorted({(f["port"],day) for f in files for day in f["timestamp_dates"]})],
            "mixed_date_rows":sum(f["mixed_date_rows"] for f in files),
            "schema_profiles":dict(Counter(f.get("schema_profile","UNVERIFIED") for f in files)),
            "blocked_files":[f["path"] for f in files if f["processing_status"]=="BLOCKED"],"blocked_months":blocked_months,
            "discovery_issues":list(discovery_issues),"ignored_files":list(ignored_files),"months":month_reports,
            "files":files,"fragment_pairs":pairs,"elapsed_seconds":time.perf_counter()-started,
            "policy":{"gps_adapter":"NOT_AUTHORIZED","mixed_date_reassignment":False,"automatic_exclusion":False,
                      "missing_days":"WARN_AND_PRESERVE_CALENDAR_BLANKS","partial_days":"OBSERVED_ENVELOPE_WARNING_ONLY",
                      "high_occupied_second_overlap_indicator_fraction":0.9,"high_time_overlap_alone_blocks":False,
                      "overlap_audit_record_limit":overlap_limit,"selector_changed":False},
            "boundary":"Source safety inspection, not execution success, research ground truth or proof of complete capture. Recognized worksheets are scanned fully; unreadable/ambiguous schemas are blocked with row coverage UNVERIFIED. No data conversion, subset exclusion or selector change."}


def write_preflight_report(report, directory):
    """Create an isolated diagnostic record; never write into source folders."""
    directory = Path(directory)
    directory.mkdir(parents=True,exist_ok=False)
    target = directory/"SOURCE_PREFLIGHT_REPORT.json"
    with target.open("x",encoding="utf-8") as stream:
        json.dump(report,stream,ensure_ascii=False,indent=2)
    lines = ["# 來源資料預檢 Source Preflight", "", f"判定：{report['status']}",
             f"實體 XLSX：{report['physical_xlsx_count']}；適用來源：{report['applicable_source_files']}；來源日：{report['source_days']}；日曆日：{report['calendar_days']}",
             f"跨日期列數：{report['mixed_date_rows']}；已掃描列數：{report['rows_scanned']}", "",
             "SAFE／SAFE_WITH_WARNINGS 只表示本次來源快照符合處理契約，不表示已產生正確成果或完整捕獲全天資料。",
             "GPS_SCHEMA_DETECTED：已知 GPS_* 欄位，時間語義尚未授權，不會自動轉換為標準時間。",
             "MIXED_DATE_ROWS：資料列日期不符檔名，整月阻擋；不自動拆檔、改日期或補缺日。",
             "SUBSET_CANDIDATE：完整 normalized core 多重集合包含的證據，不會自動排除。須另留下明確排除 manifest 與來源 hash。",
             "部分日指標只是觀測到的首末 timestamp，不等同連續捕獲或完整 24 小時；缺日保留日曆空白。", "",
             "| 港別／月份 | 檔案 | 來源日／日曆日 | 缺少來源日 | 判定 |", "|---|---:|---:|---|---|"]
    for month in report["months"]:
        lines.append(f"| {month['port']} {month['month']} | {month['physical_xlsx_count']} | {month['source_days']}/{month['calendar_days']} | {', '.join(month['missing_days']) or 'none'} | {month['status']} |")
    lines += ["", "## 被阻擋檔案"]
    lines += [f"- {f['path']}: {', '.join(f['issues']+f['blocking_group_issues'])}" for f in report['files'] if f['processing_status']=='BLOCKED']
    lines += ["", "## Fragment 證據"]
    lines += [f"- {p['left']} / {p['right']}: {', '.join(p['classifications'])}; shared core rows={p['exact_normalized_core_rows_shared']}; auto-exclusion=NO" for p in report['fragment_pairs']]
    lines += ["", "## Schema、欄位位置與部分日指標"]
    lines += [f"- {f['path']}: {f.get('schema_profile','UNVERIFIED')}; positions={f.get('field_positions_1based',{})}; timestamps={f['timestamp_dates']}; first/last second={f['min_second']}/{f['max_second']}; partial={f.get('partial_day_indicators',[])}" for f in report['files']]
    lines += ["", "## 檔案辨識問題", *[f"- {item}" for item in report['discovery_issues']], "", report['boundary']]
    (directory/"SOURCE_PREFLIGHT_SUMMARY.md").write_text("\n".join(lines)+"\n",encoding="utf-8")
    return target


def verify_source_identity(report):
    for record in report["files"]:
        path = Path(record["path"])
        stamp = path.stat()
        if (stamp.st_size,stamp.st_mtime_ns)!=(record["bytes"],record["mtime_ns"]) or sha256(path)!=record["sha256"]:
            raise ValueError("Source changed since preflight: "+str(path))
