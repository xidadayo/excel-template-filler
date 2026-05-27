from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

from src.excel_reader import read_excel_structure


def test_read_excel_structure_detects_sheets_and_fields(tmp_path: Path) -> None:
    workbook_path = tmp_path / "main.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Orders"
    sheet.append(["Order ID", "Customer", "Amount"])
    sheet.append(["SO-001", "Acme", 100])
    workbook.create_sheet("Empty")
    workbook.save(workbook_path)

    structure = read_excel_structure(workbook_path)

    assert structure["file_name"] == "main.xlsx"
    assert structure["sheet_names"] == ["Orders", "Empty"]
    orders = structure["sheets"][0]
    assert orders["max_row"] == 2
    assert orders["max_column"] == 3
    assert orders["header_row"] == 1
    assert orders["fields"] == ["Order ID", "Customer", "Amount"]
