from __future__ import annotations

import json
from pathlib import Path

from src.nl_rule_parser import apply_upload_context, parse_rules_text, write_rules_outputs


def test_parse_rules_text_generates_draft_config() -> None:
    text = """主表 Sheet「发票数据」填入 INVOICE.xlsx 模板。
客户名称填入 B5。
发票号填入 F3。
品名、数量、单价、金额填入第 10 行到第 30 行。
金额 = 数量 × 单价。
总金额写入 E31。"""

    draft, confirmations = parse_rules_text(text)

    template = draft["templates"][0]
    assert draft["need_confirm"] is True
    assert template["template_name"] == "INVOICE.xlsx"
    assert template["source_sheet"] == "发票数据"
    assert template["target_sheet"] == "Sheet1"
    assert template["mappings"][0]["source_field"] == "客户名称"
    assert template["mappings"][0]["target_cell"] == "B5"
    assert template["mappings"][1]["source_field"] == "发票号"
    assert template["detail_area"]["target_start_row"] == 10
    assert template["detail_area"]["target_end_row"] == 30
    assert template["detail_area"]["columns"]["品名"] == "A"
    assert template["detail_area"]["calculations"][0]["formula"] == "数量 * 单价"
    assert any(item["field"] == "main_file" for item in confirmations)


def test_apply_upload_context_confirms_single_uploaded_template_and_sheet() -> None:
    draft, confirmations = parse_rules_text("客户名称填入 B5。")

    apply_upload_context(
        draft,
        confirmations,
        main_file="主表.xlsx",
        template_names=["template.xlsx"],
        main_sheet_names=["订单数据"],
    )

    template = draft["templates"][0]
    assert draft["main_file"] == "主表.xlsx"
    assert template["template_name"] == "template.xlsx"
    assert template["source_sheet"] == "订单数据"
    assert draft["need_confirm"] is False


def test_write_rules_outputs(tmp_path: Path) -> None:
    draft = {"need_confirm": True, "templates": []}
    confirmations = [{"field": "main_file", "value": "", "reason": "confirm"}]
    draft_path = tmp_path / "mapping_config.draft.json"
    confirm_path = tmp_path / "mapping_config.need_confirm.json"

    write_rules_outputs(draft, confirmations, draft_path, confirm_path)

    assert json.loads(draft_path.read_text(encoding="utf-8"))["need_confirm"] is True
    assert json.loads(confirm_path.read_text(encoding="utf-8"))["need_confirm"][0]["field"] == "main_file"
