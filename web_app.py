#!/usr/bin/env python3
"""Local web app for formatting TV360 schedule text and Excel files."""

from __future__ import annotations

import json
import mimetypes
import re
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from format_tv360_schedule import (
    ScheduleRow,
    apply_corrections_to_rows,
    build_workbook_rows,
    build_xlsx_bytes,
    decode_text_bytes,
    load_corrections,
    parse_xlsx_bytes,
    parse_text,
    row_sort_key,
)


ROOT = Path(__file__).resolve().parent
STATIC_DIR = ROOT / "web"
MAX_UPLOAD_BYTES = 50 * 1024 * 1024
SAFE_FILENAME_RE = re.compile(r"[^A-Za-z0-9_. -]+")


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


class Tv360Handler(BaseHTTPRequestHandler):
    server_version = "TV360ScheduleWeb/1.0"

    def log_message(self, format: str, *args: object) -> None:
        print(f"{self.address_string()} - {format % args}")

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self.send_static("index.html")
            return
        if parsed.path in {"/app.css", "/app.js"}:
            self.send_static(parsed.path.lstrip("/"))
            return
        if parsed.path == "/favicon.ico":
            self.send_response(HTTPStatus.NO_CONTENT)
            self.end_headers()
            return
        self.send_error(HTTPStatus.NOT_FOUND, "Not found")

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path not in {"/api/preview", "/api/export"}:
            self.send_error(HTTPStatus.NOT_FOUND, "Not found")
            return

        try:
            rows, fields = self.read_uploaded_rows()
            if fields.get("sort") == "true":
                rows.sort(key=row_sort_key)

            if parsed.path == "/api/preview":
                self.send_json({"count": len(rows), "rows": rows_to_payload(rows)})
                return

            minimal = fields.get("minimal") == "true"
            query = parse_qs(parsed.query)
            output_name = safe_filename(query.get("name", ["lich_tv360.xlsx"])[0])
            if not output_name.lower().endswith(".xlsx"):
                output_name += ".xlsx"

            workbook_rows = build_workbook_rows(rows, minimal)
            payload = build_xlsx_bytes(workbook_rows, minimal)

            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Length", str(len(payload)))
            self.send_header("Content-Disposition", f'attachment; filename="{output_name}"')
            self.send_header("X-Content-Type-Options", "nosniff")
            self.end_headers()
            self.wfile.write(payload)
        except ValueError as error:
            self.send_json({"error": str(error)}, status=HTTPStatus.BAD_REQUEST)
        except Exception as error:
            self.send_json({"error": f"Lỗi xử lý: {error}"}, status=HTTPStatus.INTERNAL_SERVER_ERROR)

    def read_uploaded_rows(self) -> tuple[list[ScheduleRow], dict[str, str]]:
        content_length = int(self.headers.get("Content-Length", "0"))
        if content_length <= 0:
            raise ValueError("Chưa có dữ liệu upload.")
        if content_length > MAX_UPLOAD_BYTES:
            raise ValueError("Tổng dung lượng upload vượt quá 50MB.")

        content_type = self.headers.get("Content-Type", "")
        if not content_type.startswith("multipart/form-data"):
            raise ValueError("Request phải dùng multipart/form-data.")

        body = self.rfile.read(content_length)
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

    def send_static(self, filename: str) -> None:
        path = (STATIC_DIR / filename).resolve()
        if not str(path).startswith(str(STATIC_DIR.resolve())) or not path.is_file():
            self.send_error(HTTPStatus.NOT_FOUND, "Not found")
            return
        payload = path.read_bytes()
        content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def send_json(self, data: object, status: HTTPStatus = HTTPStatus.OK) -> None:
        payload = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)


def main() -> int:
    import argparse

    parser = argparse.ArgumentParser(description="Chạy web tool format lịch phát sóng TV360 ở local.")
    parser.add_argument("--host", default="127.0.0.1", help="Host local")
    parser.add_argument("--port", type=int, default=8765, help="Port local")
    args = parser.parse_args()

    server = ThreadingHTTPServer((args.host, args.port), Tv360Handler)
    print(f"TV360 Schedule Formatter đang chạy tại http://{args.host}:{args.port}")
    print("Nhấn Ctrl+C để dừng server.")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nĐã dừng server.")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
