from __future__ import annotations

import re
from pathlib import Path

from .models import (
    ExtractedJambLine,
    JobModel,
    PurchaseOrderLine,
    WarningMessage,
    WorkOrder,
    WorkOrderItem,
)
from .po_parser import parse_po_csv
from .units import parse_dimension_pair, parse_number, round_reasonable
from .work_order_parser import parse_work_order_xlsx


FINISH_CODES = {
    "unfinished": "NONE",
    "primed": "PRIME",
    "paint pella white": "PWHITE",
    "pella white": "PWHITE",
    "paint vinyl white": "VWHITE",
    "vinyl white": "VWHITE",
    "paint pella black": "PBLACK",
    "pella black": "PBLACK",
    "paint linen white": "LINENW",
    "linen white": "LINENW",
    "paint bright white": "BRIGHTW",
    "bright white": "BRIGHTW",
}

LUMBER_TYPES = ("Poplar", "Pine", "Red Oak", "Maple")
JAMB_THICKNESS = 0.6875


def build_job(po_path: str | Path, work_order_path: str | Path) -> JobModel:
    po_lines, po_warnings = parse_po_csv(po_path)
    work_order, wo_warnings = parse_work_order_xlsx(work_order_path)

    pella_po_number = po_lines[0].pella_po_number if po_lines else ""
    order_number = work_order.order_number or _order_from_xref(po_lines[0].pella_xref if po_lines else "")
    job = JobModel(
        po_path=Path(po_path),
        work_order_path=Path(work_order_path),
        order_number=order_number,
        pella_po_number=pella_po_number,
        customer_name=work_order.project_name or work_order.customer_name,
        jamb_lines=[],
        warnings=[*po_warnings, *wo_warnings],
    )

    for po_line in po_lines:
        work_item = work_order.items.get(po_line.xref_suffix)
        if work_item is None:
            job.warnings.append(
                WarningMessage(
                    "match.no_work_item",
                    f"No work-order item matched PO xref {po_line.pella_xref}.",
                    f"PO row {po_line.source_row}",
                )
            )
        job.jamb_lines.extend(_expand_po_line(po_line, work_item, work_order))

    return job


def _expand_po_line(
    po_line: PurchaseOrderLine, work_item: WorkOrderItem | None, work_order: WorkOrder
) -> list[ExtractedJambLine]:
    attrs = _parse_po_description(po_line.description)
    warnings: list[WarningMessage] = []
    if attrs["jamb_depth"] is None:
        warnings.append(
            WarningMessage("parse.no_depth", f"Could not parse jamb depth for {po_line.pella_xref}.")
        )
    if attrs["finish_code"] is None:
        warnings.append(
            WarningMessage("parse.no_finish", f"Could not map finish for {po_line.pella_xref}.")
        )

    units = _unit_count(attrs, po_line.quantity, warnings, po_line)
    dimensions = _dimensions_for_line(po_line, attrs, work_item, units, warnings)

    lines: list[ExtractedJambLine] = []
    for width, height, dim_warning in dimensions:
        line_warnings = list(warnings)
        if dim_warning:
            line_warnings.append(dim_warning)
        calculated_height = _calculated_jamb_height(attrs["type"], height)
        footage = _calculated_footage(attrs["type"], width, height)
        lines.append(
            ExtractedJambLine(
                po_line=po_line,
                work_item=work_item,
                type=attrs["type"],
                jamb_depth=attrs["jamb_depth"],
                jamb_width=round_reasonable(width),
                jamb_height=round_reasonable(height),
                quantity=1,
                piece_count=attrs["piece_count"],
                jamb_style=attrs["jamb_style"],
                lumber_type=attrs["lumber_type"],
                finish_text=attrs["finish_text"],
                finish_code=attrs["finish_code"],
                calculated_jamb_height=round_reasonable(calculated_height),
                calculated_footage=round_reasonable(footage),
                warnings=line_warnings,
            )
        )
    return lines


def _parse_po_description(description: str) -> dict[str, object]:
    normalized = " ".join(description.replace("–", "-").split())
    lower = normalized.lower()
    piece_count = _parse_piece_count(normalized)
    shape = "curved" if "curved" in lower else "arch" if "arch" in lower else None
    if shape == "curved":
        jamb_type = None
    elif "patio door" in lower or "3 pcs" in lower or shape == "arch":
        jamb_type = "3 Sided Patio Door"
    else:
        jamb_type = "4 Sided Window"
    jamb_style = "Drilled" if "drilled" in lower else "JE"
    lumber_type = next((material for material in LUMBER_TYPES if material.lower() in lower), None)
    finish_text = _extract_finish_text(normalized)
    finish_code = _map_finish(finish_text)

    depth_match = re.search(
        r"(?P<depth>\d+(?:\.\d+)?(?:[-\s]\d+/\d+)?|\d+/\d+)\s*\"?\s*\(Actual\)",
        normalized,
        flags=re.IGNORECASE,
    )
    jamb_depth = parse_number(depth_match.group("depth")) if depth_match else None

    return {
        "type": jamb_type,
        "jamb_depth": jamb_depth,
        "piece_count": piece_count,
        "jamb_style": jamb_style,
        "lumber_type": lumber_type,
        "finish_text": finish_text,
        "finish_code": finish_code,
        "shape": shape,
    }


def _parse_piece_count(description: str) -> int | None:
    match = re.search(r"(\d+)\s*pcs", description, flags=re.IGNORECASE)
    return int(match.group(1)) if match else None


def _extract_finish_text(description: str) -> str | None:
    star_match = re.findall(r"\*+\s*([^*]+?)\s*\*+", description)
    if star_match:
        for candidate in star_match:
            cleaned = candidate.strip().strip('"')
            if "longer" not in cleaned.lower() and cleaned:
                return cleaned

    paren_match = re.search(r"\((Unfinished|Primed)\s*\)", description, flags=re.IGNORECASE)
    if paren_match:
        return paren_match.group(1).strip()

    for key in FINISH_CODES:
        if key in description.lower():
            return key
    return None


def _map_finish(finish_text: str | None) -> str | None:
    if not finish_text:
        return None
    cleaned = " ".join(finish_text.lower().replace("*", "").strip().split())
    if cleaned in FINISH_CODES:
        return FINISH_CODES[cleaned]
    for key, code in FINISH_CODES.items():
        if key in cleaned:
            return code
    return None


def _unit_count(
    attrs: dict[str, object],
    po_quantity: int | None,
    warnings: list[WarningMessage],
    po_line: PurchaseOrderLine,
) -> int:
    piece_count = attrs["piece_count"]
    jamb_type = attrs["type"]
    if attrs.get("shape") == "curved" and isinstance(piece_count, int):
        return max(1, piece_count)
    sides = 3 if jamb_type == "3 Sided Patio Door" else 4
    if piece_count:
        units = piece_count / sides
        if units.is_integer():
            return max(1, int(units))
        warnings.append(
            WarningMessage(
                "parse.piece_count_mismatch",
                f"{po_line.pella_xref} has {piece_count} pcs, which is not divisible by {sides}.",
            )
        )
    return max(1, po_quantity or 1)


def _dimensions_for_line(
    po_line: PurchaseOrderLine,
    attrs: dict[str, object],
    work_item: WorkOrderItem | None,
    units: int,
    warnings: list[WarningMessage],
) -> list[tuple[float | None, float | None, WarningMessage | None]]:
    if work_item is None:
        return [(None, None, WarningMessage("match.no_dimensions", f"No dimensions for {po_line.pella_xref}."))]

    description = work_item.description
    lower_po = po_line.description.lower()

    if attrs.get("shape") == "curved":
        return [
            (
                None,
                None,
                WarningMessage(
                    "parse.curved_unit",
                    f"{po_line.pella_xref} is a curved unit and dimensions must be reviewed.",
                ),
            )
            for _ in range(units)
        ]

    primary = _primary_product_dimension(description)
    if primary and units == 1:
        return [(primary[0], primary[1], None)]

    components = _component_dimensions(description)
    selected = _select_components(lower_po, components, units)
    if not selected and primary:
        selected = [primary] * units
    elif not selected:
        rough = parse_dimension_pair(work_item.rough_opening or "")
        selected = [rough] * units if rough else []

    if not selected:
        return [(None, None, WarningMessage("parse.no_dimensions", f"Could not parse dimensions for {po_line.pella_xref}."))]

    results: list[tuple[float | None, float | None, WarningMessage | None]] = []
    for dimension in selected[:units]:
        warning = None
        width, height = dimension
        if "arch" in lower_po or "curved" in lower_po:
            warning = WarningMessage(
                "parse.complex_shape",
                f"{po_line.pella_xref} appears to be an arch/curved unit and should be reviewed.",
            )
        results.append((width, height, warning))
    while len(results) < units:
        results.append((None, None, WarningMessage("parse.missing_unit", f"Missing unit dimensions for {po_line.pella_xref}.")))
    return results


def _primary_product_dimension(description: str) -> tuple[float, float] | None:
    head = description[:500]
    matches = list(re.finditer(r",\s*([^,]{0,40}?\d[^,]*?\s+[Xx]\s+[^,]*?\d)\s*,", head))
    if not matches:
        return None
    return parse_dimension_pair(matches[0].group(1))


def _component_dimensions(description: str) -> list[tuple[str, float, float]]:
    matches: list[tuple[str, float, float]] = []
    pattern = re.compile(
        r"(?P<num>\d+):\s*(?P<label>.{0,140}?)Frame Size:\s*(?P<size>.{3,40}?)(?=\s+(?:Lifestyle|Architect|Impervia|Wood|Pella|General|Exterior|Interior))",
        flags=re.IGNORECASE | re.DOTALL,
    )
    for match in pattern.finditer(description):
        pair = parse_dimension_pair(match.group("size"))
        if pair:
            label = " ".join(match.group("label").split())
            matches.append((label, pair[0], pair[1]))
    return matches


def _select_components(
    po_description_lower: str, components: list[tuple[str, float, float]], units: int
) -> list[tuple[float, float]]:
    if not components:
        return []

    if "patio door" in po_description_lower:
        filtered = [(w, h) for label, w, h in components if "door" in label.lower()]
        return filtered[:units]

    if "arch" in po_description_lower or "curved" in po_description_lower:
        filtered = [(w, h) for label, w, h in components if "arch" in label.lower()]
        return filtered[:units]

    filtered = [
        (w, h)
        for label, w, h in components
        if "door" not in label.lower() and "arch" not in label.lower()
    ]
    return filtered[:units]


def _calculated_jamb_height(jamb_type: str | None, height: float | None) -> float | None:
    if height is None:
        return None
    if jamb_type == "3 Sided Patio Door":
        return height - 0.875
    return height - (JAMB_THICKNESS * 2)


def _calculated_footage(jamb_type: str | None, width: float | None, height: float | None) -> float | None:
    if width is None or height is None:
        return None
    if jamb_type == "3 Sided Patio Door":
        return (width + ((height + 2) * 2)) / 12
    return ((width * 2) + (height * 2)) / 12


def _order_from_xref(xref: str) -> str:
    return xref.split("-", 1)[0] if xref else ""
