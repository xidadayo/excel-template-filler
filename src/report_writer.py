"""Write process, validation, error, and summary reports."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook


def write_validation_report(rows: list[dict[str, Any]], path: str | Path) -> None:
    _write_table(
        rows,
        path,
        sheet_name="validation",
        headers=["check", "location", "status", "message"],
    )


def write_error_report(errors: list[dict[str, Any]], path: str | Path) -> None:
    _write_table(
        errors,
        path,
        sheet_name="errors",
        headers=["category", "location", "message"],
    )


def write_process_log(events: list[dict[str, Any]], path: str | Path) -> None:
    _write_table(
        events,
        path,
        sheet_name="process_log",
        headers=["step", "status", "message"],
    )


def write_summary(
    path: str | Path,
    main_file: str,
    templates: list[str],
    extracted_data: dict[str, Any],
    output_files: list[Path],
    errors: list[dict[str, Any]],
    validation_rows: list[dict[str, Any]],
) -> None:
    fields_by_template: dict[str, list[str]] = {}
    manual_confirm: list[str] = []
    for item in extracted_data.get("items", []):
        template = str(item.get("target_template"))
        field = str(item.get("source_field"))
        if template and field:
            fields_by_template.setdefault(template, []).append(field)
        if item.get("is_empty") or item.get("status") or item.get("is_calculated"):
            manual_confirm.append(f"{template}:{field}")
    for row in validation_rows:
        if row.get("status") != "passed":
            manual_confirm.append(f"{row.get('check')}:{row.get('location')}")

    lines = [
        "Excel template filler summary",
        f"main_file: {main_file}",
        f"templates: {', '.join(templates) if templates else ''}",
        f"success_count: {sum(1 for row in validation_rows if row.get('status') == 'passed')}",
        f"error_count: {len(errors)}",
        "",
        "Written fields by template:",
    ]
    for template, fields in fields_by_template.items():
        lines.append(f"- {template}: {', '.join(dict.fromkeys(fields))}")
    lines.extend(
        [
            "",
            "Fields/checks needing manual confirmation:",
            *(f"- {item}" for item in dict.fromkeys(manual_confirm)),
            "",
            "Output files:",
            *(f"- {path}" for path in output_files),
        ]
    )
    Path(path).write_text("\n".join(lines) + "\n", encoding="utf-8")


def _write_table(rows: list[dict[str, Any]], path: str | Path, sheet_name: str, headers: list[str]) -> None:
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    sheet.append(headers)
    for row in rows:
        sheet.append([_stringify(row.get(header, "")) for header in headers])
    workbook.save(output_path)


def _stringify(value: Any) -> Any:
    if isinstance(value, (list, tuple, set)):
        return ", ".join(str(item) for item in value)
    return value
