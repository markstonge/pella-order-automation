from __future__ import annotations

import re


_MIXED_FRACTION_RE = re.compile(
    r"^\s*(?P<whole>\d+)?(?:[-\s]+)?(?:(?P<num>\d+)\s*/\s*(?P<den>\d+))?\s*$"
)


def parse_number(value: str | int | float | None) -> float | None:
    """Parse decimals, fractions, and mixed fractions such as 3-9/16."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace('"', "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        pass

    match = _MIXED_FRACTION_RE.match(text)
    if not match:
        return None

    whole = int(match.group("whole") or 0)
    num = match.group("num")
    den = match.group("den")
    if num and den:
        denominator = int(den)
        if denominator == 0:
            return None
        return whole + int(num) / denominator
    return float(whole)


def parse_dimension_pair(text: str) -> tuple[float, float] | None:
    match = re.search(
        r"(?P<w>\d+(?:\.\d+)?(?:[-\s]\d+/\d+)?)\s*[Xx]\s*"
        r"(?P<h>\d+(?:\.\d+)?(?:[-\s]\d+/\d+)?)",
        text,
    )
    if not match:
        return None
    width = parse_number(match.group("w"))
    height = parse_number(match.group("h"))
    if width is None or height is None:
        return None
    return width, height


def round_reasonable(value: float | None) -> float | None:
    if value is None:
        return None
    return round(value, 6)
