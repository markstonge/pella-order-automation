from __future__ import annotations

import base64
import cgi
import json
import mimetypes
import tempfile
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any

from .extractor import build_job
from .workbook_writer import write_workbook


class PellaRequestHandler(SimpleHTTPRequestHandler):
    server_version = "PellaOrderAutomation/0.1"

    def do_GET(self) -> None:  # noqa: N802 - stdlib handler API
        if self.path == "/api/health":
            self._send_json({"ok": True})
            return
        super().do_GET()

    def do_POST(self) -> None:  # noqa: N802 - stdlib handler API
        if self.path != "/api/generate":
            self.send_error(404)
            return

        try:
            payload = self._handle_generate()
            self._send_json(payload)
        except Exception as exc:  # noqa: BLE001 - API boundary
            self._send_json({"ok": False, "error": str(exc)}, status=500)

    def _handle_generate(self) -> dict[str, Any]:
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={
                "REQUEST_METHOD": "POST",
                "CONTENT_TYPE": self.headers.get("Content-Type", ""),
            },
        )

        po_field = form["po"] if "po" in form else None
        work_order_field = form["work_order"] if "work_order" in form else None
        if po_field is None or work_order_field is None:
            raise ValueError("Both a purchase order CSV and work order XLSX are required.")

        with tempfile.TemporaryDirectory(prefix="pella-order-") as temp_dir:
            temp_path = Path(temp_dir)
            po_path = _save_upload(po_field, temp_path, "purchase-order.csv")
            work_order_path = _save_upload(work_order_field, temp_path, "work-order.xlsx")
            output_path = temp_path / _output_filename(work_order_field.filename)

            job = build_job(po_path, work_order_path)
            write_workbook(job, output_path)

            warnings = _dedupe_warnings(
                [*job.warnings, *(warning for line in job.jamb_lines for warning in line.warnings)]
            )
            workbook_bytes = output_path.read_bytes()

            return {
                "ok": True,
                "filename": output_path.name,
                "summary": {
                    "filename": output_path.name,
                    "order_number": job.order_number,
                    "pella_po_number": job.pella_po_number,
                    "customer_name": job.customer_name,
                    "request_dates": sorted(
                        {
                            line.po_line.requested_date.isoformat()
                            for line in job.jamb_lines
                            if line.po_line.requested_date
                        }
                    ),
                    "jamb_line_count": len(job.jamb_lines),
                    "generated_bom_rows": len(job.jamb_lines) * 4,
                },
                "warnings": [{"code": warning.code, "message": warning.message} for warning in warnings],
                "workbook_base64": base64.b64encode(workbook_bytes).decode("ascii"),
            }

    def _send_json(self, payload: dict[str, Any], status: int = 200) -> None:
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)


def run(host: str = "127.0.0.1", port: int = 8765, directory: str | Path | None = None) -> None:
    root = Path(directory or Path.cwd())
    handler = lambda *args, **kwargs: PellaRequestHandler(*args, directory=str(root), **kwargs)
    server = ThreadingHTTPServer((host, port), handler)
    print(f"Pella Order Automation API running at http://{host}:{port}")
    server.serve_forever()


def _save_upload(field, temp_path: Path, fallback_name: str) -> Path:
    filename = Path(field.filename or fallback_name).name
    suffix = Path(filename).suffix or Path(fallback_name).suffix
    path = temp_path / f"{Path(fallback_name).stem}{suffix}"
    with path.open("wb") as handle:
        handle.write(field.file.read())
    return path


def _output_filename(work_order_filename: str | None) -> str:
    stem = Path(work_order_filename or "completed-workbook").stem
    return f"{stem} - Completed.xlsx"


def _dedupe_warnings(warnings):
    unique = {}
    for warning in warnings:
        unique[(warning.code, warning.message)] = warning
    return list(unique.values())


def main() -> None:
    dist = Path.cwd() / "dist"
    directory = dist if dist.exists() else Path.cwd()
    run(directory=directory)


if __name__ == "__main__":
    main()
