from __future__ import annotations

import argparse
from pathlib import Path

from loguru import logger

from .config import load_config
from .excel_ops import run_fill
from .reports import RunReports


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Excel multi-template filler")
    parser.add_argument("--config", required=True, help="Path to JSON config file")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    config_path = Path(args.config).resolve()

    try:
        config = load_config(config_path)
        reports = RunReports(config.output_dir)
        logger.remove()
        logger.add(reports.log_path, encoding="utf-8")
        logger.add(lambda message: print(message, end=""))
        run_fill(config, reports)
        reports.write()
    except Exception as exc:
        print(f"Run failed: {exc}")
        return 1

    print(f"Run directory: {reports.run_dir}")
    return 1 if reports.issues else 0
