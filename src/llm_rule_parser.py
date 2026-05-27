"""LLM-based natural-language rule parser.

Reads natural-language rules and workbook structure files, calls a
configured LLM (OpenAI-compatible API), and produces a structured
mapping_config.draft.json plus need_confirm_questions.json.

The LLM MUST NOT modify any Excel file. It only produces configuration
drafts. Uncertain fields are marked NEED_CONFIRM_*.
"""

from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook

_HTTPX_AVAILABLE = False
try:
    import httpx  # noqa: F811

    _HTTPX_AVAILABLE = True
except ImportError:
    pass


# ---- public API ----


def parse_rules_with_llm(
    rules_text: str,
    main_structure: dict[str, Any],
    template_structure: dict[str, Any],
    llm_config: dict[str, Any],
) -> tuple[dict[str, Any], list[dict[str, str]], list[dict[str, Any]]]:
    """Run the LLM pipeline and return (draft, need_confirm, log_entries).

    Does NOT write any files — callers are responsible for saving outputs.
    """
    log_entries: list[dict[str, Any]] = []
    _add_log(log_entries, "start", "info", "LLM rule parsing started")

    system_prompt = _build_system_prompt()
    user_prompt = _build_user_prompt(rules_text, main_structure, template_structure)
    _add_log(log_entries, "prompt_built", "info", f"system={len(system_prompt)} chars, user={len(user_prompt)} chars")

    raw_response = _call_llm(system_prompt, user_prompt, llm_config, log_entries)

    draft, need_confirm = _parse_llm_response(raw_response, log_entries)

    _add_log(log_entries, "done", "info", f"draft produced: {len(draft.get('templates', []))} templates, {len(need_confirm)} need_confirm items")

    return draft, need_confirm, log_entries


def write_llm_outputs(
    draft: dict[str, Any],
    need_confirm: list[dict[str, str]],
    log_entries: list[dict[str, Any]],
    draft_path: str | Path,
    confirm_path: str | Path,
    log_path: str | Path,
) -> None:
    """Write the three output files produced by the LLM parser."""
    draft_out = Path(draft_path)
    confirm_out = Path(confirm_path)
    log_out = Path(log_path)

    draft_out.parent.mkdir(parents=True, exist_ok=True)
    confirm_out.parent.mkdir(parents=True, exist_ok=True)
    log_out.parent.mkdir(parents=True, exist_ok=True)

    draft_out.write_text(
        json.dumps(draft, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    confirm_out.write_text(
        json.dumps({"need_confirm": need_confirm}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "rule_parse_log"
    ws.append(["timestamp", "step", "level", "message"])
    for entry in log_entries:
        ws.append([entry.get(h, "") for h in ("timestamp", "step", "level", "message")])
    wb.save(log_out)


def load_llm_config(path: str | Path) -> dict[str, Any]:
    """Load and validate LLM configuration from a JSON file."""
    config_path = Path(path)
    if not config_path.exists():
        raise FileNotFoundError(f"LLM config not found: {config_path}")
    config = json.loads(config_path.read_text(encoding="utf-8"))
    required = ["model", "api_key"]
    for key in required:
        if not config.get(key):
            raise ValueError(f"LLM config missing required field: {key}")
    return config


# ---- prompt builders ----


def _build_system_prompt() -> str:
    return """You are a configuration generator for an Excel template filling system.
Your ONLY job is to convert natural-language rules into a structured JSON configuration file.

CRITICAL RULES:
1. You MUST NOT guess or invent any value. If you are unsure, use NEED_CONFIRM_<field>.
2. You ONLY output configuration files. You do NOT write Excel files or execute anything.
3. The system that calls you will validate your output before any Excel operation.
4. Every mapping must reference fields and cells that ACTUALLY EXIST in the provided structures.

OUTPUT FORMAT — you must respond with exactly this JSON structure:

```json
{
  "draft_config": {
    "project_name": "string",
    "main_file": "string (the main Excel filename)",
    "templates": [
      {
        "template_name": "string (filename of the template)",
        "output_name": "string (filename for the filled output)",
        "source_sheet": "string (sheet name in main Excel)",
        "target_sheet": "string (sheet name in template)",
        "mappings": [
          {
            "source_field": "string (field name from main sheet header row)",
            "target_cell": "string (e.g. C7, B5)",
            "write_type": "single_value",
            "required": true,
            "overwrite_formula": false
          }
        ],
        "detail_area": {
          "source_sheet": "string",
          "target_sheet": "string",
          "target_start_row": 24,
          "target_end_row": 28,
          "columns": {
            "field_name": "A",
            "another_field": "B"
          },
          "calculations": [
            {
              "target_field": "金额",
              "formula": "数量 * 单价"
            }
          ]
        }
      }
    ]
  },
  "need_confirm": [
    {
      "field": "string (path to the uncertain field)",
      "value": "string (current placeholder value)",
      "reason": "string (why it cannot be determined)"
    }
  ],
  "notes": "string (optional summary of decisions made)"
}
```

FIELD DETECTION RULES:
- source_field MUST match a field name from the main sheet's "fields" list exactly.
- target_cell MUST be a valid Excel cell reference like B5, C7, AA10.
- For detail_area, assign columns starting from the first writable column in order (A, B, C...).
- If a source_field does not exist in the main sheet fields, mark it NEED_CONFIRM_FIELD_<name>.
- If a target_cell cannot be determined, use NEED_CONFIRM_TARGET_CELL.

DETAIL AREA DETECTION:
- The detail_area represents a table/list area in the template.
- target_start_row and target_end_row define the writable row range.
- "columns" maps each source field name to the target column letter.
- A calculation formula like "金额 = 数量 * 单价" means the "金额" field is computed, not extracted from the main sheet.
- Computed fields must be listed in BOTH "columns" AND "calculations".

TEMPLATE SHEET DETECTION:
- Look at the template structure's "sheet_names" array.
- If there is only one sheet, use it as target_sheet.
- If multiple sheets exist, pick the one most likely to be an invoice/data sheet.
- If you cannot determine the correct sheet, use NEED_CONFIRM_TARGET_SHEET.

WHEN TO USE NEED_CONFIRM (not exhaustive):
1. source_sheet is ambiguous or not found in main structure
2. target_template cannot be matched to an actual template file
3. target_sheet is ambiguous or not found in template structure
4. source_field does not match any field in the main sheet header
5. target_cell cannot be inferred from the rules
6. Detail area start/end rows cannot be determined
7. Calculation formula is unclear or has syntax issues
8. Whether to allow overwrite_formula is unclear
9. Whether to allow inserting rows is unclear

IMPORTANT: The user's rules are in Chinese. Sheet names, field names, and template filenames may contain Chinese characters. Match them exactly as they appear in the provided structures."""


def _build_user_prompt(
    rules_text: str,
    main_structure: dict[str, Any],
    template_structure: dict[str, Any],
) -> str:
    main_summary = _summarize_main_structure(main_structure)
    template_summary = _summarize_template_structure(template_structure)

    return f"""=== MAIN EXCEL STRUCTURE ===
{json.dumps(main_summary, ensure_ascii=False, indent=2)}

=== TEMPLATE STRUCTURE ===
{json.dumps(template_summary, ensure_ascii=False, indent=2)}

=== USER'S NATURAL LANGUAGE RULES ===
{rules_text}

Based on the structures and rules above, generate the draft_config and need_confirm list.
Respond ONLY with the JSON object — no extra text or markdown formatting."""


def _summarize_main_structure(structure: dict[str, Any]) -> dict[str, Any]:
    return {
        "file_name": structure.get("file_name"),
        "sheet_names": structure.get("sheet_names", []),
        "sheets": [
            {
                "sheet_name": sheet["sheet_name"],
                "max_row": sheet.get("max_row"),
                "header_row": sheet.get("header_row"),
                "fields": sheet.get("fields", []),
            }
            for sheet in structure.get("sheets", [])
        ],
    }


def _summarize_template_structure(structure: dict[str, Any]) -> dict[str, Any]:
    templates = []
    for tpl in structure.get("templates", []):
        templates.append({
            "file_name": tpl.get("file_name"),
            "sheet_names": tpl.get("sheet_names", []),
            "sheets": [
                {
                    "sheet_name": sheet["sheet_name"],
                    "max_row": sheet.get("max_row"),
                    "max_column": sheet.get("max_column"),
                    "merged_cells": sheet.get("merged_cells", [])[:30],
                    "formula_cells": sheet.get("formula_cells", []),
                    "possible_detail_areas": sheet.get("possible_detail_areas", []),
                }
                for sheet in tpl.get("sheets", [])
            ],
        })
    return {"templates": templates}


# ---- LLM API call ----


def _call_llm(
    system_prompt: str,
    user_prompt: str,
    llm_config: dict[str, Any],
    log_entries: list[dict[str, Any]],
) -> str:
    if not _HTTPX_AVAILABLE:
        raise RuntimeError(
            "httpx is required for LLM calls. Install it with: pip install httpx"
        )
    base_url = llm_config.get("base_url", "https://api.openai.com/v1").rstrip("/")
    api_key = llm_config["api_key"]
    model = llm_config["model"]
    temperature = llm_config.get("temperature", 0.1)
    max_tokens = llm_config.get("max_tokens", 8192)
    timeout = llm_config.get("request_timeout", 120)

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ]

    _add_log(log_entries, "llm_call", "info", f"calling {model} at {base_url}")

    try:
        with httpx.Client(timeout=timeout) as client:
            response = client.post(
                f"{base_url}/chat/completions",
                headers={
                    "Authorization": f"Bearer {api_key}",
                    "Content-Type": "application/json",
                },
                json={
                    "model": model,
                    "messages": messages,
                    "temperature": temperature,
                    "max_tokens": max_tokens,
                },
            )
            response.raise_for_status()
    except httpx.HTTPStatusError as exc:
        detail = exc.response.text[:500] if exc.response is not None else str(exc)
        _add_log(log_entries, "llm_error", "error", f"HTTP {exc.response.status_code if exc.response else '?'}: {detail}")
        raise RuntimeError(f"LLM API returned error: {detail}") from exc
    except httpx.RequestError as exc:
        _add_log(log_entries, "llm_error", "error", f"request failed: {exc}")
        raise RuntimeError(f"LLM API request failed: {exc}") from exc

    data = response.json()
    content = ""
    choices = data.get("choices", [])
    if choices:
        content = choices[0].get("message", {}).get("content", "")
    usage = data.get("usage", {})
    _add_log(
        log_entries,
        "llm_response",
        "info",
        f"tokens: prompt={usage.get('prompt_tokens', '?')}, completion={usage.get('completion_tokens', '?')}, total={usage.get('total_tokens', '?')}",
    )

    if not content:
        _add_log(log_entries, "llm_error", "error", "LLM returned empty content")
        raise RuntimeError("LLM returned empty response content")

    return content


# ---- response parser ----


def _parse_llm_response(
    raw: str,
    log_entries: list[dict[str, Any]],
) -> tuple[dict[str, Any], list[dict[str, str]]]:
    json_text = _extract_json(raw)
    _add_log(log_entries, "parse_response", "info", f"extracted JSON block: {len(json_text)} chars")

    try:
        parsed = json.loads(json_text)
    except json.JSONDecodeError as exc:
        _add_log(log_entries, "parse_error", "error", f"JSON decode failed: {exc}")
        raise ValueError(f"LLM response is not valid JSON: {exc}") from exc

    if not isinstance(parsed, dict):
        raise ValueError("LLM response JSON is not an object")

    draft = parsed.get("draft_config", parsed.get("draft", {}))
    need_confirm = parsed.get("need_confirm", [])
    if not isinstance(need_confirm, list):
        need_confirm = []

    notes = parsed.get("notes", "")
    if notes:
        _add_log(log_entries, "llm_notes", "info", str(notes)[:500])

    if "project_name" not in draft:
        draft["project_name"] = "Excel 多模板自动填充系统"
    if "main_file" not in draft:
        draft["main_file"] = "NEED_CONFIRM_MAIN_FILE.xlsx"

    for template_cfg in draft.get("templates", []):
        if "output_name" not in template_cfg:
            tpl_name = template_cfg.get("template_name", "output")
            template_cfg["output_name"] = f"{Path(str(tpl_name)).stem}_已填充{Path(str(tpl_name)).suffix or '.xlsx'}"
        mappings = template_cfg.get("mappings")
        if isinstance(mappings, list):
            for mapping in mappings:
                mapping.setdefault("write_type", "single_value")
                mapping.setdefault("required", True)
                mapping.setdefault("overwrite_formula", False)

    _add_log(log_entries, "parse_done", "info", f"draft ready: {len(draft.get('templates', []))} template(s)")

    return draft, need_confirm


def _extract_json(text: str) -> str:
    """Extract a JSON object from LLM output that may include markdown fences."""
    text = text.strip()
    fence_match = re.search(r"```(?:json)?\s*([\s\S]*?)```", text)
    if fence_match:
        return fence_match.group(1).strip()
    brace_start = text.find("{")
    if brace_start == -1:
        return text
    depth = 0
    for i in range(brace_start, len(text)):
        ch = text[i]
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return text[brace_start : i + 1]
    return text


# ---- helpers ----


def _add_log(
    log_entries: list[dict[str, Any]],
    step: str,
    level: str,
    message: str,
) -> None:
    log_entries.append({
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "step": step,
        "level": level,
        "message": message,
    })
