"""Data cleaning helpers that preserve original and cleaned values."""

from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from typing import Any


EMPTY_MARKERS = {"", "none", "null", "nan", "n/a"}


def clean_value(value: Any) -> dict[str, Any]:
    original = value
    cleaned = _clean_scalar(value)
    return {
        "original_value": original,
        "cleaned_value": cleaned,
        "is_empty": cleaned is None,
    }


def clean_row(row: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {field: clean_value(value) for field, value in row.items()}


def _clean_scalar(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, (int, float)):
        return value

    if isinstance(value, str):
        text = " ".join(value.replace("\r", "\n").split())
        if text.lower() in EMPTY_MARKERS:
            return None
        number = _to_number(text)
        if number is not None:
            return number
        parsed_date = _to_iso_date(text)
        if parsed_date is not None:
            return parsed_date
        return text

    return value


def _to_number(text: str) -> int | float | None:
    normalized = text.replace(",", "")
    try:
        number = Decimal(normalized)
    except InvalidOperation:
        return None
    if not _looks_like_number(text):
        return None
    if number == number.to_integral_value():
        return int(number)
    return float(number)


def _looks_like_number(text: str) -> bool:
    normalized = text.replace(",", "")
    if normalized.count(".") > 1:
        return False
    if normalized.startswith("-"):
        normalized = normalized[1:]
    return bool(normalized) and all(part.isdigit() for part in normalized.split("."))


def _to_iso_date(text: str) -> str | None:
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    return None
