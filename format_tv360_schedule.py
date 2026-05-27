#!/usr/bin/env python3
"""Format raw TV360 schedule text files into an Excel workbook."""

from __future__ import annotations

import argparse
import datetime as dt
import html
import io
import re
import sys
import zipfile
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Iterable


TIME_TOKEN_RE = re.compile(
    r"""
    (?<!\d)
    (?P<hour>[01]?\d|2[0-3])
    \s*
    (?P<sep>:|h|H|\.|giờ)
    \s*
    (?P<minute>[0-5]\d)?
    (?!\d)
    """,
    re.IGNORECASE | re.VERBOSE,
)

RANGE_JOIN_RE = re.compile(r"^\s*(?:[-–—~]|to|den|đến)\s*$", re.IGNORECASE)
LEADING_MARK_RE = re.compile(r"^\s*(?:[-*•]+|\d{1,3}[\).]\s+|\d{1,3}\s*[-–]\s+)")
SPACE_RE = re.compile(r"\s+")
TITLE_TRIM_RE = re.compile(r"^[\s:;\-|–—~•*]+|[\s:;\-|–—~]+$")
EPISODE_RE = re.compile(r"\s*(?:[-–—]\s*)?\b(?P<label>tập|tap|số|so)\s+(?P<number>[0-9]+[A-Za-z]?)\b", re.IGNORECASE)
DAY_HEADER_RE = re.compile(r"^(?:thứ\s+\w+|chủ\s*nhật|cn)\s*(?:\([^)]*\))?$", re.IGNORECASE)
DATE_IN_DAY_RE = re.compile(r"\((?P<day>\d{1,2})/(?P<month>\d{1,2})(?:/(?P<year>\d{2,4}))?\)")
DATE_TEXT_RE = re.compile(r"(?P<day>\d{1,2})[/-](?P<month>\d{1,2})(?:[/-](?P<year>\d{2,4}))?")
MINUTE_ONLY_RE = re.compile(r"^(?:[0-5]?\d)$")
XML_ROW_RE = re.compile(r"<row\b[^>]*>.*?</row>", re.DOTALL)
XML_CELL_RE = re.compile(r"<c\b(?P<attrs>[^>]*)>(?P<body>.*?)</c>", re.DOTALL)
XML_ATTR_RE = re.compile(r'([A-Za-z_:][\w:.-]*)="([^"]*)"')
XML_TEXT_RE = re.compile(r"<t\b[^>]*>(.*?)</t>", re.DOTALL)
XML_VALUE_RE = re.compile(r"<v\b[^>]*>(.*?)</v>", re.DOTALL)
XML_SHARED_ITEM_RE = re.compile(r"<si\b[^>]*>(.*?)</si>", re.DOTALL)
SUPPORTED_INPUT_EXTENSIONS = {".txt", ".xlsx"}
DEFAULT_CORRECTIONS_PATH = Path(__file__).resolve().with_name("corrections.txt")
TV360_METADATA_HEADERS = [
    "Main Title",
    "Main Language",
    "Start Time",
    "End Time",
    "Main Synopsis",
    "Rating",
    "Video Type",
    "Director",
    "Actor",
    "Price",
    "Fx Point",
    "Series Key",
    "Episode Key",
    "Is Last Episode",
    "Poster Url",
    "VOD AssetID",
    "ProductID",
    "CPIP",
]
BRACKETED_TEXT_RES = (
    re.compile(r"\s*\([^()]*\)\s*"),
    re.compile(r"\s*\[[^\[\]]*\]\s*"),
    re.compile(r"\s*\{[^{}]*\}\s*"),
    re.compile(r"\s*（[^（）]*）\s*"),
    re.compile(r"\s*【[^【】]*】\s*"),
)


@dataclass(frozen=True)
class ScheduleRow:
    title: str
    airtime: str
    source: str
    source_line: int
    raw_text: str
    schedule_day: str = ""


CorrectionRules = list[tuple[str, str]]


def normalize_airtime(match: re.Match[str]) -> str:
    hour = int(match.group("hour"))
    minute_text = match.group("minute")
    minute = int(minute_text) if minute_text is not None else 0
    return f"{hour:02d}:{minute:02d}"


def airtime_sort_key(airtime: str) -> tuple[int, str]:
    first_time = airtime.split("-", 1)[0].strip()
    try:
        hour, minute = first_time.split(":", 1)
        return (int(hour) * 60 + int(minute), airtime)
    except ValueError:
        return (10_000, airtime)


def clean_text(value: str) -> str:
    value = value.replace("\ufeff", " ")
    value = value.replace("\u00a0", " ")
    return SPACE_RE.sub(" ", value).strip()


def parse_corrections_text(text: str) -> CorrectionRules:
    corrections: CorrectionRules = []
    for line_number, raw_line in enumerate(text.splitlines(), start=1):
        line = clean_text(raw_line)
        if not line or line.startswith("#"):
            continue
        if "=>" not in line:
            raise ValueError(f"Dòng corrections {line_number} thiếu dấu =>")
        wrong, correct = (part.strip() for part in line.split("=>", 1))
        if not wrong:
            raise ValueError(f"Dòng corrections {line_number} thiếu từ/cụm từ sai")
        corrections.append((wrong, correct))
    return sorted(corrections, key=lambda item: len(item[0]), reverse=True)


def load_corrections(path: Path | None = None) -> CorrectionRules:
    corrections_path = path or DEFAULT_CORRECTIONS_PATH
    if not corrections_path.is_file():
        return []
    return parse_corrections_text(read_text_file(corrections_path))


def apply_title_corrections(value: str, corrections: CorrectionRules) -> str:
    corrected = value
    for wrong, correct in corrections:
        corrected = re.sub(re.escape(wrong), correct, corrected, flags=re.IGNORECASE)
    return clean_text(corrected)


def apply_corrections_to_rows(rows: list[ScheduleRow], corrections: CorrectionRules) -> list[ScheduleRow]:
    if not corrections:
        return rows
    return [replace(row, title=apply_title_corrections(row.title, corrections)) for row in rows]


def strip_bracketed_text(value: str) -> str:
    previous = None
    current = value
    while previous != current:
        previous = current
        for pattern in BRACKETED_TEXT_RES:
            current = pattern.sub(" ", current)
    return current


def clean_title(value: str) -> str:
    value = clean_text(value)
    value = strip_bracketed_text(value)
    value = TITLE_TRIM_RE.sub("", value)
    return clean_text(value)


def normalize_episode_separator(value: str) -> str:
    def replace(match: re.Match[str]) -> str:
        label = match.group("label").lower()
        if label == "tap":
            label = "tập"
        elif label == "so":
            label = "số"
        return f" - {label} {match.group('number')}"

    return clean_text(EPISODE_RE.sub(replace, value))


def sentence_case(value: str) -> str:
    value = clean_text(value).lower()
    for index, char in enumerate(value):
        if char.isalpha():
            return value[:index] + char.upper() + value[index + 1 :]
    return value


def format_schedule_title(value: str) -> str:
    return sentence_case(normalize_episode_separator(clean_title(value)))


def strip_leading_marks(line: str) -> str:
    previous = None
    current = line
    while previous != current:
        previous = current
        current = LEADING_MARK_RE.sub("", current)
    return current.strip()


def extract_rows_from_line(line: str, source: str, source_line: int) -> list[ScheduleRow]:
    raw = clean_text(line)
    normalized = strip_leading_marks(raw)
    if not normalized:
        return []

    matches = list(TIME_TOKEN_RE.finditer(normalized))
    if not matches:
        return []

    rows: list[ScheduleRow] = []
    index = 0
    while index < len(matches):
        current = matches[index]
        prefix = format_schedule_title(normalized[: current.start()])
        title = ""
        airtime = normalize_airtime(current)
        end_pos = current.end()

        has_range = False
        if index + 1 < len(matches):
            separator = normalized[current.end() : matches[index + 1].start()]
            has_range = bool(RANGE_JOIN_RE.match(separator))

        if has_range:
            next_match = matches[index + 1]
            airtime = f"{normalize_airtime(current)}-{normalize_airtime(next_match)}"
            end_pos = next_match.end()
            next_start = matches[index + 2].start() if index + 2 < len(matches) else len(normalized)
            suffix = format_schedule_title(normalized[end_pos:next_start])
            title = suffix or prefix
            index += 2
        else:
            next_start = matches[index + 1].start() if index + 1 < len(matches) else len(normalized)
            suffix = format_schedule_title(normalized[end_pos:next_start])
            title = suffix or prefix
            index += 1

        rows.append(
            ScheduleRow(
                title=title,
                airtime=airtime,
                source=source,
                source_line=source_line,
                raw_text=raw,
            )
        )

    return rows


def read_text_file(path: Path) -> str:
    encodings = ("utf-8-sig", "utf-8", "utf-16", "cp1258", "latin-1")
    for encoding in encodings:
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def decode_text_bytes(content: bytes) -> str:
    encodings = ("utf-8-sig", "utf-8", "utf-16", "cp1258", "latin-1")
    for encoding in encodings:
        try:
            return content.decode(encoding)
        except UnicodeDecodeError:
            continue
    return content.decode("utf-8", errors="replace")


def xml_plain_text(value: str) -> str:
    return clean_text(html.unescape(re.sub(r"<[^>]+>", "", value)))


def extract_xml_text_runs(value: str) -> str:
    return clean_text("".join(xml_plain_text(match.group(1)) for match in XML_TEXT_RE.finditer(value)))


def parse_xml_attrs(value: str) -> dict[str, str]:
    return {match.group(1): html.unescape(match.group(2)) for match in XML_ATTR_RE.finditer(value)}


def excel_number_to_text(value: str) -> str:
    try:
        number = float(value)
    except ValueError:
        return value

    if 0 <= number < 1:
        total_minutes = int(round(number * 24 * 60)) % (24 * 60)
        return f"{total_minutes // 60:02d}:{total_minutes % 60:02d}"

    if number.is_integer():
        return str(int(number))
    return value


def read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        xml = archive.read("xl/sharedStrings.xml").decode("utf-8", errors="replace")
    except KeyError:
        return []
    return [extract_xml_text_runs(match.group(1)) for match in XML_SHARED_ITEM_RE.finditer(xml)]


def cell_text(cell_body: str, attrs: dict[str, str], shared_strings: list[str]) -> str:
    cell_type = attrs.get("t", "")
    if cell_type == "s":
        value_match = XML_VALUE_RE.search(cell_body)
        if not value_match:
            return ""
        try:
            return shared_strings[int(value_match.group(1))]
        except (IndexError, ValueError):
            return ""

    if cell_type == "inlineStr":
        return extract_xml_text_runs(cell_body)

    value_match = XML_VALUE_RE.search(cell_body)
    if value_match:
        value = xml_plain_text(value_match.group(1))
        if cell_type in {"str", "b", "e"}:
            return value
        return excel_number_to_text(value)

    return extract_xml_text_runs(cell_body)


def is_pure_time_cell(value: str) -> bool:
    parsed = extract_rows_from_line(value, "", 0)
    return len(parsed) == 1 and parsed[0].airtime and not parsed[0].title


def is_noise_title_cell(value: str) -> bool:
    title = clean_title(value)
    return not title or bool(re.fullmatch(r"\d{1,4}", title))


def column_index_from_ref(ref: str) -> int:
    letters = "".join(char for char in ref.upper() if "A" <= char <= "Z")
    index = 0
    for char in letters:
        index = index * 26 + ord(char) - 64
    return index


def is_day_header(value: str) -> bool:
    return bool(DAY_HEADER_RE.match(clean_text(value)))


def minute_marker_to_airtime(value: str, current_hour: int | None) -> str:
    text = clean_text(value)
    parsed = extract_rows_from_line(text, "", 0)
    if parsed and not parsed[0].title:
        return parsed[0].airtime
    if MINUTE_ONLY_RE.match(text) and current_hour is not None:
        return f"{current_hour:02d}:{int(text):02d}"
    return ""


def row_time_marker(cells: list[tuple[int, str]], day_columns: set[int], current_hour: int | None) -> str:
    left_cells = [(col, value) for col, value in cells if col < min(day_columns)]
    right_cells = [(col, value) for col, value in cells if col > max(day_columns)]
    for _col, value in left_cells + right_cells:
        airtime = minute_marker_to_airtime(value, current_hour)
        if airtime:
            return airtime
    return ""


def schedule_day_sort_key(value: str) -> tuple[int, int, str]:
    match = DATE_IN_DAY_RE.search(value)
    if match:
        year = int(match.group("year") or 0)
        month = int(match.group("month"))
        day = int(match.group("day"))
        return (year, month * 100 + day, value)
    weekdays = ["thứ hai", "thứ ba", "thứ tư", "thứ năm", "thứ sáu", "thứ bảy", "chủ nhật", "cn"]
    lower = value.lower()
    for index, weekday in enumerate(weekdays, start=1):
        if lower.startswith(weekday):
            return (0, index, value)
    return (9_999, 9_999, value)


def row_sort_key(row: ScheduleRow) -> tuple[tuple[int, int, str], tuple[int, str], str]:
    return (schedule_day_sort_key(row.schedule_day), airtime_sort_key(row.airtime), row.title)


def parse_date_text(value: str) -> dt.date | None:
    match = DATE_TEXT_RE.search(value)
    if not match:
        return None

    year_text = match.group("year")
    current_year = dt.date.today().year
    if year_text is None:
        year = current_year
    else:
        year = int(year_text)
        if year < 100:
            year += 2000

    try:
        return dt.date(year, int(match.group("month")), int(match.group("day")))
    except ValueError:
        return None


def parse_airtime_time(value: str) -> dt.time | None:
    match = re.match(r"\s*(?P<hour>\d{1,2}):(?P<minute>\d{2})\s*$", value)
    if not match:
        return None
    try:
        return dt.time(int(match.group("hour")), int(match.group("minute")))
    except ValueError:
        return None


def row_schedule_dates(records: list[ScheduleRow]) -> list[dt.date]:
    first_date = next((date for row in records if (date := parse_date_text(row.schedule_day))), None)
    fallback_date = first_date or dt.date.today()
    current_date = first_date
    dates: list[dt.date] = []
    for row in records:
        row_date = parse_date_text(row.schedule_day)
        if row_date is not None:
            current_date = row_date
        dates.append(current_date or fallback_date)
    return dates


def metadata_start_end_datetimes(records: list[ScheduleRow]) -> list[tuple[dt.datetime, dt.datetime]]:
    dates = row_schedule_dates(records)
    starts: list[dt.datetime] = []
    explicit_ends: list[dt.datetime | None] = []

    for row, schedule_date in zip(records, dates, strict=True):
        start_text, *end_parts = [part.strip() for part in row.airtime.split("-", 1)]
        start_time = parse_airtime_time(start_text) or dt.time(0, 0)
        start = dt.datetime.combine(schedule_date, start_time)
        starts.append(start)

        end = None
        if end_parts:
            end_time = parse_airtime_time(end_parts[0])
            if end_time is not None:
                end = dt.datetime.combine(schedule_date, end_time)
                if end <= start:
                    end += dt.timedelta(days=1)
        explicit_ends.append(end)

    result: list[tuple[dt.datetime, dt.datetime]] = []
    for index, start in enumerate(starts):
        if explicit_ends[index] is not None:
            end = explicit_ends[index]
        elif index + 1 < len(starts) and dates[index + 1] == dates[index] and starts[index + 1] > start:
            end = starts[index + 1]
        else:
            end = dt.datetime.combine(dates[index], dt.time(23, 59, 59))
        result.append((start, end))
    return result


def tv360_datetime_text(value: dt.datetime) -> str:
    return value.strftime("%Y%m%d%H%M%S")


def xlsx_lines_from_cells(values: list[str]) -> list[str]:
    time_cells = [(index, value) for index, value in enumerate(values) if is_pure_time_cell(value)]
    if not time_cells:
        line = clean_text(" ".join(values))
        return [line] if line else []

    lines: list[str] = []
    for position, (time_index, time_value) in enumerate(time_cells):
        next_time_index = time_cells[position + 1][0] if position + 1 < len(time_cells) else len(values)
        title_cells = values[time_index + 1 : next_time_index]

        if not title_cells and time_index > 0:
            title_cells = values[:time_index]

        for title in title_cells:
            if is_pure_time_cell(title) or is_noise_title_cell(title):
                continue
            lines.append(clean_text(f"{time_value} {title}"))

    return lines


def xlsx_rows_from_bytes(content: bytes, source: str) -> list[tuple[int, str]]:
    rows: list[tuple[int, str]] = []
    with zipfile.ZipFile(io.BytesIO(content)) as archive:
        shared_strings = read_shared_strings(archive)
        worksheet_names = sorted(
            name
            for name in archive.namelist()
            if name.startswith("xl/worksheets/") and name.endswith(".xml")
        )
        for worksheet_index, worksheet_name in enumerate(worksheet_names, start=1):
            xml = archive.read(worksheet_name).decode("utf-8", errors="replace")
            for row_index, row_match in enumerate(XML_ROW_RE.finditer(xml), start=1):
                values: list[str] = []
                for cell_match in XML_CELL_RE.finditer(row_match.group(0)):
                    attrs = parse_xml_attrs(cell_match.group("attrs"))
                    value = cell_text(cell_match.group("body"), attrs, shared_strings)
                    if value:
                        values.append(value)
                for line in xlsx_lines_from_cells(values):
                    rows.append((row_index, line))
    return rows


def worksheet_cells(xml: str, shared_strings: list[str]) -> list[tuple[int, list[tuple[int, str]]]]:
    rows: list[tuple[int, list[tuple[int, str]]]] = []
    for fallback_row_index, row_match in enumerate(XML_ROW_RE.finditer(xml), start=1):
        row_attrs = parse_xml_attrs(row_match.group(0).split(">", 1)[0])
        row_index = int(row_attrs.get("r", fallback_row_index))
        cells: list[tuple[int, str]] = []
        for cell_match in XML_CELL_RE.finditer(row_match.group(0)):
            attrs = parse_xml_attrs(cell_match.group("attrs"))
            value = cell_text(cell_match.group("body"), attrs, shared_strings)
            if not value:
                continue
            col_index = column_index_from_ref(attrs.get("r", ""))
            if col_index:
                cells.append((col_index, value))
        if cells:
            rows.append((row_index, cells))
    return rows


def xlsx_table_rows_from_bytes(content: bytes, source: str) -> list[ScheduleRow]:
    rows: list[ScheduleRow] = []
    with zipfile.ZipFile(io.BytesIO(content)) as archive:
        shared_strings = read_shared_strings(archive)
        worksheet_names = sorted(
            name
            for name in archive.namelist()
            if name.startswith("xl/worksheets/") and name.endswith(".xml")
        )
        for worksheet_name in worksheet_names:
            xml = archive.read(worksheet_name).decode("utf-8", errors="replace")
            schedule_days: dict[int, str] = {}
            current_hour: int | None = None

            for row_index, cells in worksheet_cells(xml, shared_strings):
                day_cells = [(col, value) for col, value in cells if is_day_header(value)]
                if day_cells:
                    schedule_days.update((col, clean_text(value)) for col, value in day_cells)
                    current_hour = None
                    continue

                if not schedule_days:
                    continue

                day_columns = set(schedule_days)
                airtime = row_time_marker(cells, day_columns, current_hour)
                if not airtime:
                    continue

                current_hour = int(airtime.split(":", 1)[0])
                values_by_col = {col: value for col, value in cells}
                for col in sorted(day_columns):
                    title = values_by_col.get(col, "")
                    if is_noise_title_cell(title) or is_pure_time_cell(title):
                        continue
                    rows.append(
                        ScheduleRow(
                            title=format_schedule_title(title),
                            airtime=airtime,
                            source=source,
                            source_line=row_index,
                            raw_text=clean_text(f"{airtime} {title}"),
                            schedule_day=schedule_days[col],
                        )
                    )
    return dedupe_rows(rows)


def parse_text(text: str, source: str) -> list[ScheduleRow]:
    rows: list[ScheduleRow] = []
    pending: ScheduleRow | None = None

    for line_number, line in enumerate(text.splitlines(), start=1):
        parsed = extract_rows_from_line(line, source, line_number)
        line_text = format_schedule_title(line)

        if parsed:
            pending = None
            titled_rows = [row for row in parsed if row.title]
            rows.extend(titled_rows)
            if not titled_rows and len(parsed) == 1:
                pending = parsed[0]
            continue

        if pending is not None and line_text:
            pending = ScheduleRow(
                title=line_text,
                airtime=pending.airtime,
                source=pending.source,
                source_line=pending.source_line,
                raw_text=pending.raw_text,
            )
            rows.append(pending)
            pending = None

    return rows


def dedupe_rows(rows: list[ScheduleRow]) -> list[ScheduleRow]:
    seen: set[tuple[str, str, str, int, str]] = set()
    unique_rows: list[ScheduleRow] = []
    for row in rows:
        key = (row.title, row.airtime, row.source, row.source_line, row.schedule_day)
        if key in seen:
            continue
        seen.add(key)
        unique_rows.append(row)
    return unique_rows


def parse_file(path: Path) -> list[ScheduleRow]:
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        return parse_xlsx_bytes(path.read_bytes(), path.name)
    return parse_text(read_text_file(path), path.name)


def parse_xlsx_bytes(content: bytes, source: str) -> list[ScheduleRow]:
    table_rows = xlsx_table_rows_from_bytes(content, source)
    if table_rows:
        return table_rows

    rows: list[ScheduleRow] = []
    pending: ScheduleRow | None = None

    for line_number, line in xlsx_rows_from_bytes(content, source):
        parsed = extract_rows_from_line(line, source, line_number)
        line_text = format_schedule_title(line)

        if parsed:
            pending = None
            titled_rows = [row for row in parsed if row.title]
            rows.extend(titled_rows)
            if not titled_rows and len(parsed) == 1:
                pending = parsed[0]
            continue

        if pending is not None and line_text:
            pending = ScheduleRow(
                title=line_text,
                airtime=pending.airtime,
                source=pending.source,
                source_line=pending.source_line,
                raw_text=pending.raw_text,
            )
            rows.append(pending)
            pending = None

    return dedupe_rows(rows)


def collect_input_files(inputs: Iterable[str]) -> list[Path]:
    files: list[Path] = []
    for item in inputs:
        path = Path(item).expanduser()
        if path.is_dir():
            files.extend(
                sorted(
                    p
                    for p in path.rglob("*")
                    if p.is_file() and p.suffix.lower() in SUPPORTED_INPUT_EXTENSIONS
                )
            )
        elif path.is_file():
            files.append(path)
        else:
            raise FileNotFoundError(f"Không tìm thấy file/thư mục: {item}")
    return files


def column_name(index: int) -> str:
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def xml_text(value: object) -> str:
    return html.escape(str(value), quote=False)


def make_cell(row: int, col: int, value: object, style: int = 0) -> str:
    ref = f"{column_name(col)}{row}"
    style_attr = f' s="{style}"' if style else ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}"{style_attr}><v>{value}</v></c>'
    return f'<c r="{ref}" t="inlineStr"{style_attr}><is><t>{xml_text(value)}</t></is></c>'


def make_sheet_xml(rows: list[list[object]], widths: list[int]) -> str:
    xml_rows: list[str] = []
    for row_index, values in enumerate(rows, start=1):
        cells = [
            make_cell(row_index, col_index, value, 1 if row_index == 1 else 0)
            for col_index, value in enumerate(values, start=1)
        ]
        xml_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    cols = "".join(
        f'<col min="{index}" max="{index}" width="{width}" customWidth="1"/>'
        for index, width in enumerate(widths, start=1)
    )
    last_col = column_name(len(widths))
    last_row = max(len(rows), 1)
    auto_filter = f'<autoFilter ref="A1:{last_col}{last_row}"/>' if last_row > 1 else ""

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <cols>{cols}</cols>
  <sheetData>{"".join(xml_rows)}</sheetData>
  {auto_filter}
</worksheet>"""


def xlsx_archive_files(rows: list[list[object]], minimal: bool) -> dict[str, str]:
    widths = [8, 22, 46, 22] if minimal else [36, 16, 20, 20, 46, 10, 12, 20, 20, 10, 12, 18, 18, 16, 28, 16, 16, 12]
    sheet_name = "Lich phat song" if minimal else "Metadata"
    sheet_xml = make_sheet_xml(rows, widths)
    now = dt.datetime.now(dt.UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")

    return {
        "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>""",
        "_rels/.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>""",
        "docProps/app.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>TV360 Schedule Formatter</Application>
</Properties>""",
        "docProps/core.xml": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>TV360 Schedule Formatter</dc:creator>
  <dc:title>Lịch phát sóng TV360</dc:title>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>""",
        "xl/workbook.xml": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="{sheet_name}" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>""",
        "xl/_rels/workbook.xml.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>""",
        "xl/styles.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>""",
        "xl/worksheets/sheet1.xml": sheet_xml,
    }


def build_xlsx_bytes(rows: list[list[object]], minimal: bool) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, content in xlsx_archive_files(rows, minimal).items():
            archive.writestr(name, content)
    return buffer.getvalue()


def write_xlsx(path: Path, rows: list[list[object]], minimal: bool) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(build_xlsx_bytes(rows, minimal))


def build_workbook_rows(records: list[ScheduleRow], minimal: bool) -> list[list[object]]:
    if minimal:
        rows: list[list[object]] = [["STT", "Thứ ngày", "Tiêu đề chương trình", "Thời gian phát sóng"]]
        rows.extend([[index, row.schedule_day, row.title, row.airtime] for index, row in enumerate(records, start=1)])
        return rows

    rows = [TV360_METADATA_HEADERS]
    rows.extend([[""] * len(TV360_METADATA_HEADERS) for _ in range(17)])
    for row, (start, end) in zip(records, metadata_start_end_datetimes(records), strict=True):
        rows.append(
            [
                row.title,
                "vie",
                tv360_datetime_text(start),
                tv360_datetime_text(end),
                row.title,
                "0",
                "HD",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
            ]
        )
    return rows


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Chuẩn hóa file text lịch phát sóng TV360 thành file Excel.",
    )
    parser.add_argument("inputs", nargs="+", help="Một hoặc nhiều file .txt/.xlsx, hoặc thư mục chứa các file đó")
    parser.add_argument("-o", "--output", default="output/tv360_schedule.xlsx", help="Đường dẫn file .xlsx đầu ra")
    parser.add_argument("--sort", action="store_true", help="Sắp xếp kết quả theo thời gian phát sóng")
    parser.add_argument("--minimal", action="store_true", help="Chỉ xuất các cột chính: STT, thứ ngày, tiêu đề, thời gian")
    parser.add_argument(
        "--corrections",
        default=str(DEFAULT_CORRECTIONS_PATH),
        help="File từ điển sửa chính tả dạng 'sai => đúng'. Mặc định dùng corrections.txt nếu có",
    )
    parser.add_argument("--dry-run", action="store_true", help="In kết quả parse ra màn hình, không tạo Excel")
    return parser.parse_args(argv)


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    try:
        files = collect_input_files(args.inputs)
        if not files:
            print("Không có file .txt nào để xử lý.", file=sys.stderr)
            return 2

        records: list[ScheduleRow] = []
        for file_path in files:
            records.extend(parse_file(file_path))

        records = apply_corrections_to_rows(records, load_corrections(Path(args.corrections).expanduser()))

        if args.sort:
            records.sort(key=row_sort_key)

        if args.dry_run:
            for row in records:
                day_text = f"{row.schedule_day}\t" if row.schedule_day else ""
                print(f"{day_text}{row.airtime}\t{row.title}\t{row.source}:{row.source_line}")
            print(f"Tổng số chương trình parse được: {len(records)}")
            return 0

        rows = build_workbook_rows(records, args.minimal)
        output_path = Path(args.output).expanduser()
        write_xlsx(output_path, rows, args.minimal)
        print(f"Đã tạo {output_path} với {len(records)} chương trình từ {len(files)} file.")
        return 0
    except Exception as error:
        print(f"Lỗi: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
