"""Generate a fill plan from mapping config and workbook structures."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.utils.cell import coordinate_to_tuple

from src.rule_parser import MappingConfig, TemplateMappingConfig


FILL_PLAN_HEADERS = [
    "source_file",
    "source_sheet",
    "source_field",
    "target_template",
    "output_name",
    "target_sheet",
    "target_cell",
    "write_type",
    "is_calculated",
    "calculation_formula",
    "required",
    "overwrite_formula",
    "risk_tips",
]

ERROR_HEADERS = ["category", "location", "message"]


def generate_fill_plan(
    config: MappingConfig | None,
    main_structure: dict[str, Any] | None,
    template_structure: dict[str, Any] | None,
    input_main_file: Path | None,
    template_files: list[Path],
    template_aliases: dict[str, Path] | None = None,
) -> tuple[list[dict[str, Any]], list[dict[str, str]]]:
    if config is None:
        return [], []

    errors: list[dict[str, str]] = []
    plan: list[dict[str, Any]] = []
    main_sheets = _sheets_by_name(main_structure)
    templates_by_name = {path.name: path for path in template_files}
    aliases = template_aliases or {}
    templates_by_name.update(aliases)
    template_structures = {
        template["file_name"]: template
        for template in (template_structure or {}).get("templates", [])
    }
    for alias_name, converted_path in aliases.items():
        if converted_path.name in template_structures:
            template_structures[alias_name] = template_structures[converted_path.name]

    if input_main_file and config.main_file != input_main_file.name:
        errors.append(
            _error(
                "config",
                "main_file",
                f"configured main_file '{config.main_file}' does not match detected file '{input_main_file.name}'",
            )
        )

    if not config.templates:
        errors.append(_error("config", "templates", "at least one template mapping is required"))

    for template_config in config.templates:
        _append_template_plan(
            config,
            template_config,
            main_sheets,
            templates_by_name,
            template_structures,
            plan,
            errors,
        )

    return plan, errors


def write_fill_plan_json(plan: list[dict[str, Any]], path: str | Path) -> None:
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps({"fill_plan": plan}, ensure_ascii=False, indent=2), encoding="utf-8")


def write_fill_plan_xlsx(plan: list[dict[str, Any]], path: str | Path) -> None:
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "fill_plan"
    sheet.append(FILL_PLAN_HEADERS)
    for item in plan:
        sheet.append([_stringify(item.get(header, "")) for header in FILL_PLAN_HEADERS])
    workbook.save(output_path)


def write_config_error_report(errors: list[dict[str, str]], path: str | Path) -> None:
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "config_errors"
    sheet.append(ERROR_HEADERS)
    for error in errors:
        sheet.append([error.get(header, "") for header in ERROR_HEADERS])
    workbook.save(output_path)


def _append_template_plan(
    config: MappingConfig,
    template_config: TemplateMappingConfig,
    main_sheets: dict[str, dict[str, Any]],
    templates_by_name: dict[str, Path],
    template_structures: dict[str, dict[str, Any]],
    plan: list[dict[str, Any]],
    errors: list[dict[str, str]],
) -> None:
    template_name = template_config.template_name
    template_structure = template_structures.get(template_name)
    template_sheets = _sheets_by_name(template_structure)
    source_sheet = main_sheets.get(template_config.source_sheet)
    target_sheet = template_sheets.get(template_config.target_sheet)

    if template_name not in templates_by_name:
        errors.append(_error("template", template_name, "target_template does not exist in input/templates"))
    if source_sheet is None:
        errors.append(_error("source_sheet", template_config.source_sheet, "source_sheet does not exist in main Excel"))
    if template_structure is not None and target_sheet is None:
        errors.append(_error("target_sheet", f"{template_name}:{template_config.target_sheet}", "target_sheet does not exist"))

    source_fields = set((source_sheet or {}).get("fields", []))
    formula_cells = {cell["cell"] for cell in (target_sheet or {}).get("formula_cells", [])}

    for mapping in template_config.mappings:
        location = f"{template_name}:{template_config.target_sheet}!{mapping.target_cell}"
        risks: list[str] = []
        if not _is_valid_cell(mapping.target_cell):
            errors.append(_error("target_cell", location, "target_cell is not a valid Excel cell reference"))
            risks.append("目标单元格地址不合法")
        if source_sheet is not None and mapping.source_field not in source_fields:
            errors.append(_error("source_field", mapping.source_field, "source_field does not exist in source_sheet fields"))
            risks.append("来源字段不存在")
        if target_sheet is not None and mapping.target_cell in formula_cells and not mapping.overwrite_formula:
            risks.append("目标单元格包含公式，默认禁止覆盖")

        plan.append(
            {
                "source_file": config.main_file,
                "source_sheet": template_config.source_sheet,
                "source_field": mapping.source_field,
                "target_template": template_name,
                "output_name": template_config.output_name,
                "target_sheet": template_config.target_sheet,
                "target_cell": mapping.target_cell,
                "write_type": mapping.write_type,
                "is_calculated": False,
                "calculation_formula": "",
                "required": mapping.required,
                "overwrite_formula": mapping.overwrite_formula,
                "risk_tips": risks,
            }
        )

    if template_config.detail_area is not None:
        _append_detail_plan(config, template_config, main_sheets, template_sheets, plan, errors)


def _append_detail_plan(
    config: MappingConfig,
    template_config: TemplateMappingConfig,
    main_sheets: dict[str, dict[str, Any]],
    template_sheets: dict[str, dict[str, Any]],
    plan: list[dict[str, Any]],
    errors: list[dict[str, str]],
) -> None:
    detail = template_config.detail_area
    if detail is None:
        return

    source_sheet = main_sheets.get(detail.source_sheet)
    target_sheet = template_sheets.get(detail.target_sheet)
    if source_sheet is None:
        errors.append(_error("source_sheet", detail.source_sheet, "detail_area source_sheet does not exist in main Excel"))
    if target_sheet is None and template_sheets:
        errors.append(_error("target_sheet", f"{template_config.template_name}:{detail.target_sheet}", "detail_area target_sheet does not exist"))
    if detail.target_end_row < detail.target_start_row:
        errors.append(_error("detail_area", "target_end_row", "target_end_row must be greater than or equal to target_start_row"))

    source_fields = set((source_sheet or {}).get("fields", []))
    calculations = {calculation.target_field: calculation.formula for calculation in detail.calculations}

    for source_field, target_column in detail.columns.items():
        risks: list[str] = []
        if not _is_valid_detail_column(target_column):
            errors.append(_error("detail_area.columns", source_field, f"target column '{target_column}' is not valid"))
            risks.append("目标列不合法")
        if source_sheet is not None and source_field not in source_fields and source_field not in calculations:
            errors.append(_error("source_field", source_field, "detail source_field does not exist in source_sheet fields"))
            risks.append("来源字段不存在")

        plan.append(
            {
                "source_file": config.main_file,
                "source_sheet": detail.source_sheet,
                "source_field": source_field,
                "target_template": template_config.template_name,
                "output_name": template_config.output_name,
                "target_sheet": detail.target_sheet,
                "target_cell": f"{target_column}{detail.target_start_row}",
                "target_range": f"{target_column}{detail.target_start_row}:{target_column}{detail.target_end_row}",
                "write_type": "detail_column",
                "is_calculated": source_field in calculations,
                "calculation_formula": calculations.get(source_field, ""),
                "required": True,
                "overwrite_formula": source_field in calculations,
                "risk_tips": risks,
            }
        )


def _sheets_by_name(structure: dict[str, Any] | None) -> dict[str, dict[str, Any]]:
    if not structure:
        return {}
    return {sheet["sheet_name"]: sheet for sheet in structure.get("sheets", [])}


def _is_valid_cell(value: str) -> bool:
    try:
        row, column = coordinate_to_tuple(value)
        return row >= 1 and column >= 1
    except ValueError:
        return False


def _is_valid_detail_column(value: str) -> bool:
    return value.isalpha() and value.upper() == value


def _stringify(value: Any) -> Any:
    if isinstance(value, list):
        return "; ".join(str(item) for item in value)
    return value


def _error(category: str, location: str, message: str) -> dict[str, str]:
    return {"category": category, "location": location, "message": message}
