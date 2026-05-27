"""Parse natural-language rules into a mapping_config draft.

The parser is deterministic: when a value is uncertain it writes a
NEED_CONFIRM placeholder instead of guessing.
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any


# Use ASCII-only word boundaries so Chinese/Unicode characters don't block
# cell-reference detection (Python 3 treats CJK chars as \w, which defeats \b).
CELL_RE = re.compile(r"(?<![A-Za-z0-9_])([A-Z]{1,3}[1-9][0-9]*)(?![A-Za-z0-9_])")
SHEET_RE = re.compile(r"(?:主表\s*)?Sheet\s*[「“\"]([^」”\"]+)[」”\"]|工作表\s*[「“\"]([^」”\"]+)[」”\"]")
TEMPLATE_RE = re.compile(r"([A-Za-z0-9_\- .（）()]+?\.(?:xlsx|xlsm|xls))", re.IGNORECASE)
ROW_RANGE_RE = re.compile(r"第\s*(\d+)\s*行\s*(?:到|至|-|~)\s*第?\s*(\d+)\s*行")
CALC_RE = re.compile(r"^([^=＝]+?)\s*[=＝]\s*(.+)$")


def parse_rules_file(path: str | Path) -> tuple[dict[str, Any], list[dict[str, str]]]:
    rules_path = Path(path)
    text = rules_path.read_text(encoding="utf-8")
    return parse_rules_text(text)


def parse_rules_text(text: str) -> tuple[dict[str, Any], list[dict[str, str]]]:
    lines = [_normalize_line(line) for line in text.splitlines()]
    lines = [line for line in lines if line]
    confirmations: list[dict[str, str]] = []

    source_sheet = _extract_source_sheet(lines)
    target_template = _extract_template(lines)
    target_sheet = _extract_target_sheet(lines)
    mappings = _extract_single_mappings(lines)
    detail_area = _extract_detail_area(lines, source_sheet, target_sheet, confirmations)

    if source_sheet is None:
        confirmations.append(_confirm("source_sheet", "", "无法识别主表 Sheet，请补充例如：主表 Sheet「发票数据」。"))
        source_sheet = "NEED_CONFIRM_SOURCE_SHEET"
    if target_template is None:
        confirmations.append(_confirm("target_template", "", "无法识别模板文件名，请补充例如：填入 INVOICE.xlsx 模板。"))
        target_template = "NEED_CONFIRM_TEMPLATE.xlsx"
    if target_sheet is None:
        confirmations.append(_confirm("target_sheet", "", "无法识别目标 Sheet，已暂用 Sheet1。"))
        target_sheet = "Sheet1"

    if not mappings and detail_area is None:
        confirmations.append(_confirm("mappings", "", "没有识别到字段映射，请补充例如：客户名称填入 B5。"))

    for mapping in mappings:
        if str(mapping["target_cell"]).startswith("NEED_CONFIRM"):
            confirmations.append(_confirm("target_cell", mapping["source_field"], "无法识别目标单元格。"))

    output_name = _default_output_name(target_template)
    draft = {
        "project_name": "Excel 多模板自动填充系统",
        "main_file": "NEED_CONFIRM_MAIN_FILE.xlsx",
        "need_confirm": True,
        "templates": [
            {
                "template_name": target_template,
                "output_name": output_name,
                "source_sheet": source_sheet,
                "target_sheet": target_sheet,
                "mappings": mappings,
            }
        ],
    }
    if detail_area is not None:
        draft["templates"][0]["detail_area"] = detail_area

    confirmations.append(_confirm("main_file", "", "主表文件名需由上传文件确认。"))
    return draft, confirmations


def apply_upload_context(
    draft: dict[str, Any],
    confirmations: list[dict[str, str]],
    *,
    main_file: str,
    template_names: list[str],
    main_sheet_names: list[str] | None = None,
    template_sheets_map: dict[str, list[str]] | None = None,
    rules_text: str | None = None,
) -> list[dict[str, str]]:
    """Fill safe values from uploaded files and keep unresolved items visible."""

    draft["main_file"] = main_file
    _remove_confirmation(confirmations, "main_file")
    templates = draft.get("templates") or []
    if not templates:
        return confirmations

    template = templates[0]
    current_template = str(template.get("template_name", ""))
    if current_template.startswith("NEED_CONFIRM") and len(template_names) == 1:
        uploaded_template = template_names[0]
        template["template_name"] = uploaded_template
        template["output_name"] = _default_output_name(uploaded_template)
        _replace_confirmation(
            confirmations,
            "target_template",
            _confirm("target_template", uploaded_template, "已根据唯一上传模板自动确认。"),
        )

    current_sheet = str(template.get("source_sheet", ""))
    if current_sheet.startswith("NEED_CONFIRM") and main_sheet_names:
        _auto_confirm_source_sheet(template, main_sheet_names, rules_text, confirmations)

    resolved_template_name = str(template.get("template_name", ""))
    current_target_sheet = str(template.get("target_sheet", "Sheet1"))
    if template_sheets_map:
        matched_sheets = template_sheets_map.get(resolved_template_name)
        if matched_sheets is None:
            for name in template_names:
                if name == resolved_template_name:
                    matched_sheets = template_sheets_map.get(name)
                    break
    else:
        matched_sheets = None
    if matched_sheets is not None:
        template_sheet_names = matched_sheets
        if template_sheet_names and _needs_target_sheet_fix(current_target_sheet, template_sheet_names):
            if len(template_sheet_names) == 1:
                template["target_sheet"] = template_sheet_names[0]
                if isinstance(template.get("detail_area"), dict):
                    template["detail_area"]["target_sheet"] = template_sheet_names[0]
                _replace_confirmation(
                    confirmations,
                    "target_sheet",
                    _confirm("target_sheet", template_sheet_names[0], f"已根据模板唯一 Sheet「{template_sheet_names[0]}」自动确认。"),
                )
            else:
                picked = _pick_most_likely_target_sheet(template_sheet_names, current_target_sheet)
                template["target_sheet"] = picked
                if isinstance(template.get("detail_area"), dict):
                    template["detail_area"]["target_sheet"] = picked
                _replace_confirmation(
                    confirmations,
                    "target_sheet",
                    _confirm("target_sheet", picked, f"模板「{resolved_template_name}」含多个 Sheet，已选用「{picked}」，如需更换请在规则中指定。"),
                )
                template["_target_sheet_options"] = template_sheet_names

    unresolved = _has_unresolved_placeholders(draft)
    draft["need_confirm"] = unresolved
    return confirmations


def _auto_confirm_source_sheet(
    template: dict[str, Any],
    main_sheet_names: list[str],
    rules_text: str | None,
    confirmations: list[dict[str, str]],
) -> None:
    if len(main_sheet_names) == 1:
        template["source_sheet"] = main_sheet_names[0]
        if isinstance(template.get("detail_area"), dict):
            template["detail_area"]["source_sheet"] = main_sheet_names[0]
        _replace_confirmation(
            confirmations,
            "source_sheet",
            _confirm("source_sheet", main_sheet_names[0], "已根据主表唯一 Sheet 自动确认。"),
        )
        return

    if rules_text:
        matches = _match_sheet_names_from_text(rules_text, main_sheet_names)
        if len(matches) == 1:
            template["source_sheet"] = matches[0]
            if isinstance(template.get("detail_area"), dict):
                template["detail_area"]["source_sheet"] = matches[0]
            _replace_confirmation(
                confirmations,
                "source_sheet",
                _confirm("source_sheet", matches[0], f"已根据规则文本匹配主表 Sheet「{matches[0]}」。"),
            )
            return
        if len(matches) > 1:
            pass


def _match_sheet_names_from_text(rules_text: str, sheet_names: list[str]) -> list[str]:
    matched: list[str] = []
    for name in sheet_names:
        if name in rules_text:
            matched.append(name)
    if matched:
        return matched
    cleaned_text = rules_text.replace(" ", "").replace("　", "")
    for name in sheet_names:
        cleaned_name = name.replace(" ", "").replace("　", "")
        if cleaned_name and cleaned_name in cleaned_text:
            matched.append(name)
    return matched


def _needs_target_sheet_fix(current: str, template_sheet_names: list[str]) -> bool:
    if current.startswith("NEED_CONFIRM"):
        return True
    if current not in template_sheet_names:
        return True
    return False


def _pick_most_likely_target_sheet(sheet_names: list[str], current: str) -> str:
    for name in sheet_names:
        if name.lower() in ("sheet1", "arkusz1", "tabela1", "feuil1", "tabella1"):
            return name
    non_empty = [name for name in sheet_names if name and not name.startswith("_")]
    if non_empty:
        return non_empty[0]
    return sheet_names[0]


def write_rules_outputs(
    draft: dict[str, Any],
    confirmations: list[dict[str, str]],
    draft_path: str | Path,
    confirmation_path: str | Path,
) -> None:
    draft_output = Path(draft_path)
    confirm_output = Path(confirmation_path)
    draft_output.parent.mkdir(parents=True, exist_ok=True)
    confirm_output.parent.mkdir(parents=True, exist_ok=True)
    draft_output.write_text(json.dumps(draft, ensure_ascii=False, indent=2), encoding="utf-8")
    confirm_output.write_text(json.dumps({"need_confirm": confirmations}, ensure_ascii=False, indent=2), encoding="utf-8")


def _extract_source_sheet(lines: list[str]) -> str | None:
    for line in lines:
        match = SHEET_RE.search(line)
        if match:
            return next(group for group in match.groups() if group)
    return None


def _extract_template(lines: list[str]) -> str | None:
    for line in lines:
        match = TEMPLATE_RE.search(line)
        if match:
            return match.group(1).strip()
    return None


def _extract_target_sheet(lines: list[str]) -> str | None:
    for line in lines:
        if "目标" in line or "模板 Sheet" in line or "target_sheet" in line:
            match = SHEET_RE.search(line)
            if match:
                return next(group for group in match.groups() if group)
    return None


def _extract_single_mappings(lines: list[str]) -> list[dict[str, Any]]:
    mappings: list[dict[str, Any]] = []
    for line in lines:
        if "=" in line or "＝" in line or _looks_like_detail_line(line):
            continue
        # Strip template filenames so cell-like substrings inside them
        # (e.g. "SM260501" in "INVOICE bank transfer SM260501.xlsx")
        # are not mistaken for cell references.
        clean_line = TEMPLATE_RE.sub("", line)
        cell_match = CELL_RE.search(clean_line)
        if not cell_match:
            continue
        # Derive source field from the original line text before the match,
        # using the cleaned-line offset mapped back into the original line.
        field_text = clean_line[: cell_match.start()]
        field = re.sub(r"(填入|写入|放到|到|至|->|→)", " ", field_text).strip(" ：:，,。；;")
        if not field:
            field = "NEED_CONFIRM_FIELD"
        mappings.append(
            {
                "source_field": field,
                "target_cell": cell_match.group(1),
                "write_type": "single_value",
                "required": True,
                "overwrite_formula": False,
            }
        )
    return mappings


def _extract_detail_area(
    lines: list[str],
    source_sheet: str | None,
    target_sheet: str | None,
    confirmations: list[dict[str, str]],
) -> dict[str, Any] | None:
    detail_line = next((line for line in lines if _looks_like_detail_line(line)), None)
    if detail_line is None:
        return None

    row_match = ROW_RANGE_RE.search(detail_line)
    if row_match:
        start_row = int(row_match.group(1))
        end_row = int(row_match.group(2))
    else:
        start_row = 1
        end_row = 1
        confirmations.append(_confirm("detail_area.rows", detail_line, "无法识别明细起止行。"))

    fields_text = re.split(r"(?:填入|写入|放到)", detail_line, maxsplit=1)[0]
    fields = [field.strip() for field in re.split(r"[、，,]", fields_text) if field.strip()]
    columns = {field: _column_for_offset(offset) for offset, field in enumerate(fields)}

    calculations = []
    for line in lines:
        calc_match = CALC_RE.match(line)
        if not calc_match:
            continue
        target_field = calc_match.group(1).strip()
        expression = (
            calc_match.group(2)
            .strip()
            .replace("×", "*")
            .replace("＊", "*")
            .replace("x", "*")
            .replace("X", "*")
        )
        calculations.append({"target_field": target_field, "formula": expression})
        if target_field not in columns:
            columns[target_field] = _column_for_offset(len(columns))

    if not fields:
        confirmations.append(_confirm("detail_area.columns", detail_line, "无法识别明细字段。"))

    return {
        "source_sheet": source_sheet or "NEED_CONFIRM_SOURCE_SHEET",
        "target_sheet": target_sheet or "Sheet1",
        "target_start_row": start_row,
        "target_end_row": end_row,
        "columns": columns,
        "calculations": calculations,
        "need_confirm": True,
    }


def _looks_like_detail_line(line: str) -> bool:
    return bool(ROW_RANGE_RE.search(line)) and any(separator in line for separator in ("、", "，", ","))


def _column_for_offset(offset: int) -> str:
    result = ""
    index = offset
    while True:
        index, remainder = divmod(index, 26)
        result = chr(ord("A") + remainder) + result
        if index == 0:
            return result
        index -= 1


def _default_output_name(template_name: str) -> str:
    path = Path(template_name)
    suffix = path.suffix or ".xlsx"
    return f"{path.stem}_已填充{suffix}"


def _normalize_line(line: str) -> str:
    return line.strip().strip("。；;")


def _confirm(field: str, value: str, reason: str) -> dict[str, str]:
    return {"field": field, "value": value, "reason": reason}


def _remove_confirmation(confirmations: list[dict[str, str]], field: str) -> None:
    confirmations[:] = [item for item in confirmations if item.get("field") != field]


def _replace_confirmation(confirmations: list[dict[str, str]], field: str, replacement: dict[str, str]) -> None:
    _remove_confirmation(confirmations, field)
    confirmations.append(replacement)


def _has_unresolved_placeholders(value: Any) -> bool:
    if isinstance(value, str):
        return value.startswith("NEED_CONFIRM")
    if isinstance(value, dict):
        return any(_has_unresolved_placeholders(item) for item in value.values())
    if isinstance(value, list):
        return any(_has_unresolved_placeholders(item) for item in value)
    return False
