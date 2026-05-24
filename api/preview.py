from __future__ import annotations

from http import HTTPStatus
from http.server import BaseHTTPRequestHandler

from api._common import read_uploaded_rows, rows_to_payload, send_json
from format_tv360_schedule import row_sort_key


class handler(BaseHTTPRequestHandler):
    def do_POST(self) -> None:
        try:
            rows, fields = read_uploaded_rows(self)
            if fields.get("sort") == "true":
                rows.sort(key=row_sort_key)
            send_json(self, {"count": len(rows), "rows": rows_to_payload(rows)})
        except ValueError as error:
            send_json(self, {"error": str(error)}, status=HTTPStatus.BAD_REQUEST)
        except Exception as error:
            send_json(self, {"error": f"Lỗi xử lý: {error}"}, status=HTTPStatus.INTERNAL_SERVER_ERROR)

    def do_GET(self) -> None:
        send_json(self, {"error": "Method not allowed."}, status=HTTPStatus.METHOD_NOT_ALLOWED)
