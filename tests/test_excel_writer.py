from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

from src.template_analyzer import analyze_template


def test_analyze_template_detects_formula_merged_and_non_empty_cells(tmp_path: Path) -> None:
    template_path = tmp_path / "template.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Template"
    sheet.merge_cells("A1:B1")
    sheet["A1"] = "Invoice"
    sheet["A3"] = "SKU"
    sheet["B3"] = "Qty"
    sheet["C10"] = "=SUM(B4:B9)"
    workbook.save(template_path)

    structure = analyze_template(template_path)

    template_sheet = structure["sheets"][0]
    assert "A1:B1" in template_sheet["merged_cells"]
    assert {"cell": "C10", "formula": "=SUM(B4:B9)"} in template_sheet["formula_cells"]
    assert {"cell": "A1", "value_type": "str"} in template_sheet["non_empty_cells"]
    assert template_sheet["possible_detail_areas"]
