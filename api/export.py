from __future__ import annotations

import tempfile
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from api._common import read_uploaded_rows, safe_filename, send_json
from format_tv360_schedule import build_workbook_rows, row_sort_key, write_xlsx


class handler(BaseHTTPRequestHandler):
    def do_POST(self) -> None:
        try:
            rows, fields = read_uploaded_rows(self)
            if fields.get("sort") == "true":
                rows.sort(key=row_sort_key)

            minimal = fields.get("minimal") == "true"
            query = parse_qs(urlparse(self.path).query)
            output_name = safe_filename(query.get("name", ["lich_tv360.xlsx"])[0])
            if not output_name.lower().endswith(".xlsx"):
                output_name += ".xlsx"

            workbook_rows = build_workbook_rows(rows, minimal)
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=True) as temp_file:
                write_xlsx(Path(temp_file.name), workbook_rows, minimal)
                temp_file.seek(0)
                payload = temp_file.read()

            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Length", str(len(payload)))
            self.send_header("Content-Disposition", f'attachment; filename="{output_name}"')
            self.end_headers()
            self.wfile.write(payload)
        except ValueError as error:
            send_json(self, {"error": str(error)}, status=HTTPStatus.BAD_REQUEST)
        except Exception as error:
            send_json(self, {"error": f"Lỗi xử lý: {error}"}, status=HTTPStatus.INTERNAL_SERVER_ERROR)

    def do_GET(self) -> None:
        send_json(self, {"error": "Method not allowed."}, status=HTTPStatus.METHOD_NOT_ALLOWED)
