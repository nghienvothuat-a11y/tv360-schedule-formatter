from __future__ import annotations

import json
import re
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler
from pathlib import Path

from format_tv360_schedule import (
    ScheduleRow,
    apply_corrections_to_rows,
    decode_text_bytes,
    load_corrections,
    parse_text,
    parse_xlsx_bytes,
)


MAX_UPLOAD_BYTES = 50 * 1024 * 1024
SAFE_FILENAME_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]+')


def safe_filename(value: str) -> str:
    name = SAFE_FILENAME_RE.sub("_", value).strip(" ._")
    return name or "lich_tv360.xlsx"


def parse_content_disposition(value: str) -> dict[str, str]:
    parts = [part.strip() for part in value.split(";")]
    result: dict[str, str] = {}
    for part in parts[1:]:
        if "=" not in part:
            continue
        key, raw_value = part.split("=", 1)
        result[key.strip().lower()] = raw_value.strip().strip('"')
    return result


def parse_multipart(content_type: str, body: bytes) -> tuple[list[dict[str, object]], dict[str, str]]:
    match = re.search(r"boundary=(?P<boundary>[^;]+)", content_type)
    if not match:
        raise ValueError("Thiếu multipart boundary.")

    boundary = match.group("boundary").strip().strip('"').encode("utf-8")
    delimiter = b"--" + boundary
    files: list[dict[str, object]] = []
    fields: dict[str, str] = {}

    for raw_part in body.split(delimiter):
        part = raw_part.strip()
        if not part or part == b"--":
            continue
        if part.endswith(b"--"):
            part = part[:-2].rstrip()
        if b"\r\n\r\n" not in part:
            continue

        header_bytes, content = part.split(b"\r\n\r\n", 1)
        content = content.rstrip(b"\r\n")
        headers: dict[str, str] = {}
        for line in header_bytes.decode("utf-8", errors="replace").split("\r\n"):
            if ":" not in line:
                continue
            key, value = line.split(":", 1)
            headers[key.strip().lower()] = value.strip()

        disposition = parse_content_disposition(headers.get("content-disposition", ""))
        field_name = disposition.get("name", "")
        filename = disposition.get("filename")
        if filename:
            files.append({"field": field_name, "filename": Path(filename).name, "content": content})
        elif field_name:
            fields[field_name] = decode_text_bytes(content)

    return files, fields


def read_uploaded_rows(handler: BaseHTTPRequestHandler) -> tuple[list[ScheduleRow], dict[str, str]]:
    content_length = int(handler.headers.get("Content-Length", "0"))
    if content_length <= 0:
        raise ValueError("Chưa có dữ liệu upload.")
    if content_length > MAX_UPLOAD_BYTES:
        raise ValueError("Tổng dung lượng upload vượt quá 50MB.")

    content_type = handler.headers.get("Content-Type", "")
    if not content_type.startswith("multipart/form-data"):
        raise ValueError("Request phải dùng multipart/form-data.")

    body = handler.rfile.read(content_length)
    files, fields = parse_multipart(content_type, body)
    input_files = [
        item
        for item in files
        if str(item.get("filename", "")).lower().endswith((".txt", ".xlsx"))
    ]
    if not input_files:
        raise ValueError("Hãy chọn ít nhất một file .txt hoặc .xlsx.")

    rows: list[ScheduleRow] = []
    for item in input_files:
        filename = str(item["filename"])
        content = item["content"]
        if not isinstance(content, bytes):
            continue
        if filename.lower().endswith(".xlsx"):
            rows.extend(parse_xlsx_bytes(content, filename))
        else:
            rows.extend(parse_text(decode_text_bytes(content), filename))
    if fields.get("corrections") == "true":
        rows = apply_corrections_to_rows(rows, load_corrections())
    return rows, fields


def rows_to_payload(rows: list[ScheduleRow]) -> list[dict[str, object]]:
    return [
        {
            "stt": index,
            "title": row.title,
            "airtime": row.airtime,
            "schedule_day": row.schedule_day,
            "source": row.source,
            "source_line": row.source_line,
            "raw_text": row.raw_text,
        }
        for index, row in enumerate(rows, start=1)
    ]


def send_json(handler: BaseHTTPRequestHandler, data: object, status: HTTPStatus = HTTPStatus.OK) -> None:
    payload = json.dumps(data, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(payload)))
    handler.end_headers()
    handler.wfile.write(payload)
