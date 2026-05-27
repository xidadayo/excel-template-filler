"""Simple calculation support for extracted Excel data."""

from __future__ import annotations

import ast
from pathlib import Path
from typing import Any

from openpyxl import Workbook


CALCULATION_LOG_HEADERS = [
    "plan_index",
    "source_sheet",
    "target_field",
    "source_row",
    "formula",
    "input_values",
    "result",
    "status",
    "message",
]


def apply_calculations(extracted_data: dict[str, Any]) -> tuple[dict[str, Any], list[dict[str, Any]], list[dict[str, str]]]:
    errors: list[dict[str, str]] = []
    logs: list[dict[str, Any]] = []
    sheet_rows = extracted_data.get("sheet_rows", {})

    for item in extracted_data.get("items", []):
        if not item.get("is_calculated"):
            continue
        source_sheet = item.get("source_sheet")
        source_field = item.get("source_field")
        formula = item.get("calculation_formula", "")
        rows = sheet_rows.get(source_sheet, {}).get("rows", [])
        calculated_rows = []

        for row in rows:
            context = row.get("cleaned_values", {})
            try:
                result = calculate_expression(formula, context)
                calculated_rows.append(
                    {
                        "source_row": row["row_number"],
                        "raw_value": result,
                        "cleaned_value": result,
                        "is_empty": result is None,
                    }
                )
                logs.append(_log(item, row["row_number"], formula, context, result, "success", ""))
            except Exception as exc:
                message = str(exc)
                calculated_rows.append(
                    {
                        "source_row": row["row_number"],
                        "raw_value": None,
                        "cleaned_value": None,
                        "is_empty": True,
                        "error": message,
                    }
                )
                logs.append(_log(item, row["row_number"], formula, context, "", "failed", message))
                errors.append(_error("calculation", f"{source_sheet}.{source_field}.row{row['row_number']}", message))

        item["rows"] = calculated_rows

    return extracted_data, logs, errors


def calculate_expression(formula: str, context: dict[str, Any]) -> Any:
    expression = formula.strip()
    if not expression:
        raise ValueError("empty calculation formula")
    if expression.lower().startswith("sum(") and expression.endswith(")"):
        field = expression[4:-1].strip()
        value = context.get(field)
        if isinstance(value, list):
            return sum(_number(item, field) for item in value)
        return _number(value, field)
    tree = ast.parse(expression, mode="eval")
    return _eval_node(tree.body, context)


def write_calculation_log(logs: list[dict[str, Any]], path: str | Path) -> None:
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "calculation_log"
    sheet.append(CALCULATION_LOG_HEADERS)
    for log in logs:
        sheet.append([log.get(header, "") for header in CALCULATION_LOG_HEADERS])
    workbook.save(output_path)


def write_data_error_report(errors: list[dict[str, str]], path: str | Path) -> None:
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "data_errors"
    sheet.append(["category", "location", "message"])
    for error in errors:
        sheet.append([error.get("category", ""), error.get("location", ""), error.get("message", "")])
    workbook.save(output_path)


def _eval_node(node: ast.AST, context: dict[str, Any]) -> Any:
    if isinstance(node, ast.BinOp):
        left = _eval_node(node.left, context)
        right = _eval_node(node.right, context)
        if isinstance(node.op, ast.Add):
            return left + right
        if isinstance(node.op, ast.Sub):
            return left - right
        if isinstance(node.op, ast.Mult):
            return left * right
        if isinstance(node.op, ast.Div):
            return left / right
        raise ValueError(f"unsupported operator: {type(node.op).__name__}")
    if isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.USub):
        return -_eval_node(node.operand, context)
    if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
        return node.value
    if isinstance(node, ast.Name):
        return _number(context.get(node.id), node.id)
    raise ValueError(f"unsupported expression node: {type(node).__name__}")


def _number(value: Any, field_name: str) -> int | float:
    if value is None:
        raise ValueError(f"field '{field_name}' is empty")
    if isinstance(value, bool):
        raise ValueError(f"field '{field_name}' is not numeric")
    if isinstance(value, (int, float)):
        return value
    try:
        return float(str(value).replace(",", ""))
    except ValueError as exc:
        raise ValueError(f"field '{field_name}' is not numeric: {value}") from exc


def _log(
    item: dict[str, Any],
    source_row: int,
    formula: str,
    context: dict[str, Any],
    result: Any,
    status: str,
    message: str,
) -> dict[str, Any]:
    return {
        "plan_index": item.get("plan_index"),
        "source_sheet": item.get("source_sheet"),
        "target_field": item.get("source_field"),
        "source_row": source_row,
        "formula": formula,
        "input_values": str(context),
        "result": result,
        "status": status,
        "message": message,
    }


def _error(category: str, location: str, message: str) -> dict[str, str]:
    return {"category": category, "location": location, "message": message}
