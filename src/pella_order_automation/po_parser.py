from __future__ import annotations

import csv
import re
from datetime import date, datetime
from pathlib import Path

from .models import PurchaseOrderLine, WarningMessage


_XREF_RE = re.compile(r"(?P<order>[A-Z0-9]+)-(?P<suffix>\d{3,})")


def parse_po_csv(path: str | Path) -> tuple[list[PurchaseOrderLine], list[WarningMessage]]:
    path = Path(path)
    warnings: list[WarningMessage] = []
    lines: list[PurchaseOrderLine] = []

    if path.suffix.lower() != ".csv":
        raise ValueError(f"Version 1 only supports CSV purchase orders: {path.name}")

    with path.open(newline="", encoding="utf-8-sig") as handle:
        rows = list(csv.reader(handle))

    for row_number, row in enumerate(rows, start=1):
        if not row:
            continue
        xref_index = _find_xref_index(row)
        if xref_index is None:
            continue

        xref = row[xref_index].strip()
        match = _XREF_RE.fullmatch(xref)
        if not match:
            continue

        try:
            suffix = int(match.group("suffix"))
        except ValueError:
            warnings.append(
                WarningMessage("po.bad_xref", f"Could not parse PO xref suffix from {xref}.", str(path))
            )
            continue

        po_number = _value_after(row, "PO Number") or ""
        po_line_number = row[xref_index - 1].strip() if xref_index > 0 else ""
        description = row[xref_index + 1].strip() if xref_index + 1 < len(row) else ""
        requested_date = _parse_short_date(row[xref_index + 2].strip() if xref_index + 2 < len(row) else "")
        quantity = _parse_int(row[xref_index + 3].strip() if xref_index + 3 < len(row) else "")

        lines.append(
            PurchaseOrderLine(
                source_row=row_number,
                po_line_number=po_line_number,
                pella_xref=xref,
                xref_suffix=suffix,
                description=description,
                requested_date=requested_date,
                quantity=quantity,
                pella_po_number=po_number,
            )
        )

    if not lines:
        warnings.append(WarningMessage("po.no_lines", "No purchase-order lines were found.", str(path)))

    return lines, warnings


def _find_xref_index(row: list[str]) -> int | None:
    for index, value in enumerate(row):
        if _XREF_RE.fullmatch(value.strip()):
            return index
    return None


def _value_after(row: list[str], label: str) -> str | None:
    for index, value in enumerate(row):
        if value.strip().lower() == label.lower() and index + 1 < len(row):
            return row[index + 1].strip()
    return None


def _parse_int(text: str) -> int | None:
    try:
        return int(float(text))
    except (TypeError, ValueError):
        return None


def _parse_short_date(text: str) -> date | None:
    if not text:
        return None
    for fmt in ("%d-%b", "%d-%B", "%m/%d/%Y"):
        try:
            parsed = datetime.strptime(text, fmt)
            year = parsed.year if "%Y" in fmt else 2026
            return date(year, parsed.month, parsed.day)
        except ValueError:
            continue
    return None
