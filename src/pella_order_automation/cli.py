from __future__ import annotations

import argparse
import json
from dataclasses import asdict, is_dataclass
from datetime import date
from pathlib import Path
from typing import Any

from .extractor import build_job
from .workbook_writer import write_workbook


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="pella-orders")
    subparsers = parser.add_subparsers(dest="command", required=True)

    generate = subparsers.add_parser("generate", help="Generate a completed workbook.")
    generate.add_argument("--po", required=True, help="Path to CSV purchase order.")
    generate.add_argument("--work-order", required=True, help="Path to XLSX work order.")
    generate.add_argument("--output", required=True, help="Output workbook path.")
    generate.add_argument("--json", action="store_true", help="Print normalized job JSON.")
    generate.add_argument("--summary-json", action="store_true", help="Print machine-readable generation summary only.")

    args = parser.parse_args(argv)
    if args.command == "generate":
        job = build_job(args.po, args.work_order)
        output = write_workbook(job, args.output)
        warnings = _dedupe_warnings([*job.warnings, *(warning for line in job.jamb_lines for warning in line.warnings)])
        if args.summary_json:
            print(json.dumps(_summary_payload(job, output, warnings)))
            return 0

        print(f"Generated: {output}")
        if warnings:
            print("Warnings:")
            for warning in warnings:
                print(f"- [{warning.code}] {warning.message}")
        if args.json:
            print(json.dumps(_to_jsonable(job), indent=2))
        return 0

    return 1


def _summary_payload(job, output: Path, warnings) -> dict[str, Any]:
    return {
        "ok": True,
        "filename": output.name,
        "summary": {
            "filename": output.name,
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
    }


def _dedupe_warnings(warnings):
    unique = {}
    for warning in warnings:
        unique[(warning.code, warning.message)] = warning
    return list(unique.values())


def _to_jsonable(value: Any) -> Any:
    if is_dataclass(value):
        return {key: _to_jsonable(item) for key, item in asdict(value).items()}
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, list):
        return [_to_jsonable(item) for item in value]
    if isinstance(value, dict):
        return {key: _to_jsonable(item) for key, item in value.items()}
    return value


if __name__ == "__main__":
    raise SystemExit(main())
