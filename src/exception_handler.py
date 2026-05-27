"""Report helpers for fill-time exceptions and summaries."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook


ERROR_HEADERS = ["category", "location", "message"]


def write_error_report(errors: list[dict[str, str]], path: str | Path) -> None:
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "errors"
    sheet.append(ERROR_HEADERS)
    for error in errors:
        sheet.append([error.get(header, "") for header in ERROR_HEADERS])
    workbook.save(output_path)


def write_summary(path: str | Path, output_files: list[Path], errors: list[dict[str, str]]) -> None:
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    lines = [
        "Excel template filler summary",
        f"filled_files: {len(output_files)}",
        f"errors: {len(errors)}",
    ]
    if output_files:
        lines.append("")
        lines.append("Generated files:")
        lines.extend(str(file_path) for file_path in output_files)
    if errors:
        lines.append("")
        lines.append("Warnings and errors:")
        lines.extend(f"[{item.get('category')}] {item.get('location')}: {item.get('message')}" for item in errors)
    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def error(category: str, location: str, message: str) -> dict[str, str]:
    return {"category": category, "location": location, "message": message}
