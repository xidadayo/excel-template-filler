from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook

from src.exception_handler import write_error_report
from src.rule_parser import load_mapping_config


ROOT = Path(__file__).resolve().parents[1]


def test_readme_contains_required_delivery_sections() -> None:
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    required = [
        "项目介绍",
        "安装方法",
        "目录说明",
        "如何放置主表",
        "如何放置模板",
        "如何配置 mapping_config.json",
        "如何运行分析",
        "如何生成填充计划",
        "如何执行填充",
        "如何查看报告",
        "常见错误说明",
    ]
    for section in required:
        assert section in readme


def test_example_mapping_config_is_parseable() -> None:
    config, errors = load_mapping_config(ROOT / "config" / "mapping_config.example.json")

    assert errors == []
    assert config is not None
    assert config.templates[0].mappings[0].source_field == "客户名称"


def test_final_commands_are_listed_in_help() -> None:
    result = subprocess.run(
        [
            sys.executable,
            "main.py",
            "--help",
        ],
        cwd=ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        check=True,
    )

    for command in ("analyze", "plan", "extract", "fill", "run"):
        assert command in result.stdout


def test_error_report_can_be_generated(tmp_path: Path) -> None:
    report = tmp_path / "error_report.xlsx"
    write_error_report(
        [{"category": "source_field", "location": "Sheet1.客户名称", "message": "missing"}],
        report,
    )

    sheet = load_workbook(report).active
    assert sheet["A2"].value == "source_field"
    assert sheet["C2"].value == "missing"
