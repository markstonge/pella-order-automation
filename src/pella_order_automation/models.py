from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from pathlib import Path


@dataclass
class WarningMessage:
    code: str
    message: str
    source: str | None = None


@dataclass
class PurchaseOrderLine:
    source_row: int
    po_line_number: str
    pella_xref: str
    xref_suffix: int
    description: str
    requested_date: date | None
    quantity: int | None
    pella_po_number: str


@dataclass
class WorkOrderItem:
    source_row: int
    item_number: int
    quantity: int | None
    rough_opening: str | None
    location: str | None
    comment: str | None
    description: str


@dataclass
class WorkOrder:
    path: Path
    order_number: str
    customer_name: str
    project_name: str
    items: dict[int, WorkOrderItem]


@dataclass
class ExtractedJambLine:
    po_line: PurchaseOrderLine
    work_item: WorkOrderItem | None
    type: str | None
    jamb_depth: float | None
    jamb_width: float | None
    jamb_height: float | None
    quantity: int
    piece_count: int | None
    jamb_style: str | None
    lumber_type: str | None
    finish_text: str | None
    finish_code: str | None
    calculated_jamb_height: float | None
    calculated_footage: float | None
    warnings: list[WarningMessage] = field(default_factory=list)


@dataclass
class JobModel:
    po_path: Path
    work_order_path: Path
    order_number: str
    pella_po_number: str
    customer_name: str
    ship_city: str = "Green Bay"
    ship_state: str = "WI"
    ship_address: str = "Pella"
    ship_zip: str = "54201"
    schedule: str | None = None
    date_code: str | None = None
    jamb_lines: list[ExtractedJambLine] = field(default_factory=list)
    warnings: list[WarningMessage] = field(default_factory=list)
