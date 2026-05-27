from __future__ import annotations

import argparse
import json
from pathlib import Path

from src.data_calculator import apply_calculations, write_calculation_log, write_data_error_report
from src.data_extractor import extract_data_from_plan
from src.config_validator import validate_draft_config, write_validation_report
from src.exception_handler import write_error_report as write_basic_error_report
from src.exception_handler import write_summary as write_basic_summary
from src.excel_reader import read_excel_structure
from src.excel_writer import write_filled_templates
from src.file_manager import FileManager
from src.fill_plan_generator import (
    generate_fill_plan,
    write_config_error_report,
    write_fill_plan_json,
    write_fill_plan_xlsx,
)
from src.nl_rule_parser import parse_rules_file, write_rules_outputs
from src.report_writer import (
    write_error_report,
    write_process_log,
    write_summary,
    write_validation_report,
)
from src.rule_parser import load_mapping_config
from src.template_analyzer import analyze_templates
from src.validator import validate_outputs


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-template-filler",
        description="Excel multi-template filler command line entry point.",
    )
    parser.add_argument(
        "--mapping-config",
        default="config/mapping_config.example.json",
        help="Path to the mapping JSON config file.",
    )
    parser.add_argument(
        "--system-config",
        default="config/system_config.example.json",
        help="Path to the system JSON config file.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Validate command line inputs without filling Excel files.",
    )
    subparsers = parser.add_subparsers(dest="command")
    subparsers.add_parser("analyze", help="Analyze main Excel and template structures.")
    plan_parser = subparsers.add_parser("plan", help="Generate fill_plan.json and fill_plan.xlsx from mapping config.")
    plan_parser.add_argument(
        "--config",
        default="config/mapping_config.example.json",
        help="Path to mapping_config.json.",
    )
    extract_parser = subparsers.add_parser("extract", help="Extract, clean, and calculate data from the main Excel file.")
    extract_parser.add_argument(
        "--config",
        default="config/mapping_config.example.json",
        help="Path to mapping_config.json.",
    )
    fill_parser = subparsers.add_parser("fill", help="Safely fill copied Excel templates.")
    fill_parser.add_argument(
        "--config",
        default="config/mapping_config.example.json",
        help="Path to mapping_config.json.",
    )
    run_parser = subparsers.add_parser("run", help="Run the full analyze/plan/extract/fill/validate workflow.")
    run_parser.add_argument(
        "--config",
        default="config/mapping_config.example.json",
        help="Path to mapping_config.json.",
    )
    parse_rules_parser = subparsers.add_parser("parse-rules", help="Parse natural-language rules into a mapping config draft.")
    parse_rules_parser.add_argument(
        "--rules",
        default="config/rules_prompt.txt",
        help="Path to natural-language rules prompt file.",
    )
    parse_rules_parser.add_argument(
        "--llm",
        action="store_true",
        help="Use LLM (large language model) to parse rules instead of the deterministic parser.",
    )
    parse_rules_parser.add_argument(
        "--llm-config",
        default="config/llm_config.json",
        help="Path to LLM configuration file (required when --llm is set).",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    mapping_config = Path(args.mapping_config)
    system_config = Path(args.system_config)

    if args.command == "analyze":
        return run_analyze()
    if args.command == "plan":
        return run_plan(Path(args.config))
    if args.command == "extract":
        return run_extract(Path(args.config))
    if args.command == "fill":
        return run_fill(Path(args.config))
    if args.command == "run":
        return run_all(Path(args.config))
    if args.command == "parse-rules":
        if getattr(args, "llm", False):
            return run_parse_rules_llm(Path(args.rules), Path(args.llm_config))
        return run_parse_rules(Path(args.rules))

    print("excel-template-filler")
    print(f"mapping_config: {mapping_config}")
    print(f"system_config: {system_config}")
    print(f"dry_run: {args.dry_run}")
    print("Use `python main.py analyze`, `python main.py plan`, `python main.py extract`, `python main.py fill`, `python main.py run`, or `python main.py parse-rules`.")
    return 0


def run_parse_rules(rules_path: Path) -> int:
    draft, confirmations = parse_rules_file(rules_path)
    draft_path = Path("config/mapping_config.draft.json")
    confirmation_path = Path("config/mapping_config.need_confirm.json")
    write_rules_outputs(draft, confirmations, draft_path, confirmation_path)
    print(f"Draft mapping config: {draft_path}")
    print(f"Need-confirm list: {confirmation_path}")
    print("No Excel files were modified. Review and confirm the draft before using it as mapping_config.json.")
    return 0


def run_parse_rules_llm(rules_path: Path, llm_config_path: Path) -> int:
    from src.llm_rule_parser import (
        load_llm_config,
        parse_rules_with_llm,
        write_llm_outputs,
    )

    if not rules_path.exists():
        print(f"Error: rules file not found: {rules_path}")
        return 1
    if not llm_config_path.exists():
        print(f"Error: LLM config not found: {llm_config_path}")
        print("Copy config/llm_config.example.json to config/llm_config.json and fill in your API key.")
        return 1

    rules_text = rules_path.read_text(encoding="utf-8")
    if not rules_text.strip():
        print(f"Error: rules file is empty: {rules_path}")
        return 1

    file_manager = FileManager()
    task_paths = file_manager.create_task_dirs()

    print("Step 1/4: Reading workbook structures ...")
    try:
        main_structure_json = task_paths.logs_dir / "main_excel_structure.json"
        template_structure_json = task_paths.logs_dir / "template_structure.json"
        if not main_structure_json.exists() or not template_structure_json.exists():
            print("Structure files not found. Running analyze step first ...")
            run_analyze()
        main_structure = json.loads(main_structure_json.read_text(encoding="utf-8"))
        template_structure = json.loads(template_structure_json.read_text(encoding="utf-8"))
        print(f"  Main file: {main_structure.get('file_name')} ({len(main_structure.get('sheets', []))} sheets)")
        print(f"  Templates: {len(template_structure.get('templates', []))} file(s)")
    except Exception as exc:
        print(f"Error loading structure files: {exc}")
        print("Run `python main.py analyze` first to generate structure data.")
        return 1

    print("Step 2/4: Loading LLM config ...")
    try:
        llm_config = load_llm_config(llm_config_path)
        print(f"  Provider: {llm_config.get('provider', 'openai')}")
        print(f"  Model: {llm_config['model']}")
    except Exception as exc:
        print(f"Error loading LLM config: {exc}")
        return 1

    print("Step 3/4: Calling LLM to parse rules ...")
    try:
        draft, need_confirm, log_entries = parse_rules_with_llm(
            rules_text=rules_text,
            main_structure=main_structure,
            template_structure=template_structure,
            llm_config=llm_config,
        )
        print(f"  Draft generated: {len(draft.get('templates', []))} template(s)")
        print(f"  Need-confirm items: {len(need_confirm)}")
    except Exception as exc:
        print(f"Error calling LLM: {exc}")
        return 1

    print("Step 4/4: Validating and writing outputs ...")
    draft_path = Path("config/mapping_config.draft.json")
    confirm_path = Path("config/mapping_config.need_confirm.json")
    log_path = task_paths.logs_dir / "rule_parse_log.xlsx"

    write_llm_outputs(draft, need_confirm, log_entries, draft_path, confirm_path, log_path)

    passed, nc, failed = validate_draft_config(
        draft,
        main_structure=main_structure,
        template_structure=template_structure,
    )
    validation_path = task_paths.reports_dir / "config_validation_report.xlsx"
    write_validation_report(passed, nc, failed, validation_path)

    print(f"  Draft config: {draft_path}")
    print(f"  Need-confirm list: {confirm_path}")
    print(f"  Parse log: {log_path}")
    print(f"  Validation report: {validation_path}")
    print(f"  Validation: {len(passed)} passed, {len(nc)} need_confirm, {len(failed)} failed")

    if failed:
        print()
        print("WARNING: Validation found issues that must be fixed before using this config:")
        for item in failed[:10]:
            print(f"  - [{item.get('check')}] {item.get('location')}: {item.get('message')}")
        if len(failed) > 10:
            print(f"  ... and {len(failed) - 10} more issues")
        print()
        print("Review the validation report and fix the issues before using the config.")

    if need_confirm or _has_need_confirm_draft(draft):
        print()
        print("NOTICE: Some fields could not be determined.")
        print("Review config/mapping_config.need_confirm.json and resolve before filling.")
        print("No Excel files were modified.")

    print("No Excel files were modified. Review and confirm the draft before using it as mapping_config.json.")
    return 1 if failed else 0


def _has_need_confirm_draft(draft: dict[str, object]) -> bool:
    from src.config_validator import has_unresolved_confirmations
    return has_unresolved_confirmations(draft)


def run_analyze() -> int:
    file_manager = FileManager()
    task_paths = file_manager.create_task_dirs()
    inputs = file_manager.collect_inputs()

    main_structure = read_excel_structure(inputs.main_excel)
    template_structure = analyze_templates(inputs.templates)

    main_output = task_paths.logs_dir / "main_excel_structure.json"
    template_output = task_paths.logs_dir / "template_structure.json"
    main_output.write_text(json.dumps(main_structure, ensure_ascii=False, indent=2), encoding="utf-8")
    template_output.write_text(json.dumps(template_structure, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Main Excel analyzed: {inputs.main_excel}")
    print(f"Templates analyzed: {len(inputs.templates)}")
    print(f"Main structure JSON: {main_output}")
    print(f"Template structure JSON: {template_output}")
    return 0


def run_plan(config_path: Path) -> int:
    file_manager = FileManager()
    task_paths = file_manager.create_task_dirs()
    config, errors = load_mapping_config(config_path)
    inputs = None
    main_structure = None
    template_structure = None

    try:
        inputs = file_manager.collect_inputs()
    except Exception as exc:
        errors.append({"category": "input", "location": "input", "message": str(exc)})

    if inputs is not None:
        try:
            main_structure = read_excel_structure(inputs.main_excel)
            template_structure = analyze_templates(inputs.templates)
        except Exception as exc:
            errors.append({"category": "excel", "location": "input files", "message": str(exc)})

    plan, plan_errors = generate_fill_plan(
        config=config,
        main_structure=main_structure,
        template_structure=template_structure,
        input_main_file=inputs.main_excel if inputs else None,
        template_files=inputs.templates if inputs else [],
    )
    errors.extend(plan_errors)

    fill_plan_json = task_paths.logs_dir / "fill_plan.json"
    fill_plan_xlsx = task_paths.logs_dir / "fill_plan.xlsx"
    config_error_report = task_paths.reports_dir / "config_error_report.xlsx"
    write_fill_plan_json(plan, fill_plan_json)
    write_fill_plan_xlsx(plan, fill_plan_xlsx)
    write_config_error_report(errors, config_error_report)

    print(f"Fill plan rows: {len(plan)}")
    print(f"Config errors: {len(errors)}")
    print(f"Fill plan JSON: {fill_plan_json}")
    print(f"Fill plan XLSX: {fill_plan_xlsx}")
    print(f"Config error report: {config_error_report}")
    return 1 if errors else 0


def run_extract(config_path: Path) -> int:
    file_manager = FileManager()
    task_paths = file_manager.create_task_dirs()
    plan, errors, inputs = _build_plan_for_command(config_path, task_paths)

    extracted_data = {"source_file": "", "items": [], "sheet_rows": {}}
    calculation_logs = []
    if inputs is not None and plan:
        extracted_data, extract_errors = extract_data_from_plan(inputs.main_excel, plan)
        extracted_data, calculation_logs, calculation_errors = apply_calculations(extracted_data)
        errors.extend(extract_errors)
        errors.extend(calculation_errors)
    elif inputs is None:
        errors.append({"category": "input", "location": "input", "message": "data extraction skipped because inputs are invalid"})

    extracted_output = task_paths.logs_dir / "extracted_data.json"
    calculation_log_output = task_paths.logs_dir / "calculation_log.xlsx"
    data_error_output = task_paths.reports_dir / "data_error_report.xlsx"
    extracted_output.write_text(json.dumps(extracted_data, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    write_calculation_log(calculation_logs, calculation_log_output)
    write_data_error_report(errors, data_error_output)

    print(f"Extracted items: {len(extracted_data.get('items', []))}")
    print(f"Calculation log rows: {len(calculation_logs)}")
    print(f"Data errors: {len(errors)}")
    print(f"Extracted data JSON: {extracted_output}")
    print(f"Calculation log XLSX: {calculation_log_output}")
    print(f"Data error report: {data_error_output}")
    return 1 if errors else 0


def run_fill(config_path: Path) -> int:
    file_manager = FileManager()
    task_paths = file_manager.create_task_dirs()
    plan, errors, inputs = _build_plan_for_command(config_path, task_paths)

    extracted_data = {"source_file": "", "items": [], "sheet_rows": {}}
    calculation_logs = []
    output_files = []
    if inputs is not None and plan:
        extracted_data, extract_errors = extract_data_from_plan(inputs.main_excel, plan)
        extracted_data, calculation_logs, calculation_errors = apply_calculations(extracted_data)
        errors.extend(extract_errors)
        errors.extend(calculation_errors)
        output_files, write_errors = write_filled_templates(
            extracted_data=extracted_data,
            template_files=inputs.templates,
            output_dir=task_paths.filled_files_dir,
            allow_insert_rows=False,
            template_aliases=inputs.template_aliases,
        )
        errors.extend(write_errors)
    elif inputs is None:
        errors.append({"category": "input", "location": "input", "message": "fill skipped because inputs are invalid"})

    extracted_output = task_paths.logs_dir / "extracted_data.json"
    calculation_log_output = task_paths.logs_dir / "calculation_log.xlsx"
    error_report_output = task_paths.reports_dir / "error_report.xlsx"
    data_error_output = task_paths.reports_dir / "data_error_report.xlsx"
    extracted_output.write_text(json.dumps(extracted_data, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    write_calculation_log(calculation_logs, calculation_log_output)
    write_data_error_report(errors, data_error_output)
    write_basic_error_report(errors, error_report_output)
    write_basic_summary(task_paths.summary_file, output_files, errors)

    print(f"Filled files: {len(output_files)}")
    print(f"Errors: {len(errors)}")
    print(f"Filled output dir: {task_paths.filled_files_dir}")
    print(f"Error report: {error_report_output}")
    print(f"Summary: {task_paths.summary_file}")
    return 1 if errors else 0


def run_all(config_path: Path) -> int:
    file_manager = FileManager()
    task_paths = file_manager.create_task_dirs()
    process_events: list[dict[str, str]] = []
    errors: list[dict[str, str]] = []
    validation_rows: list[dict[str, object]] = []
    extracted_data = {"source_file": "", "items": [], "sheet_rows": {}}
    output_files: list[Path] = []
    inputs = None

    try:
        plan, plan_errors, inputs = _build_plan_for_command(config_path, task_paths)
        errors.extend(plan_errors)
        process_events.append({"step": "plan", "status": "done", "message": f"fill plan rows={len(plan)}"})
    except Exception as exc:
        plan = []
        errors.append({"category": "plan", "location": str(config_path), "message": str(exc)})
        process_events.append({"step": "plan", "status": "failed", "message": str(exc)})

    try:
        if inputs is None:
            raise ValueError("input files are invalid")
        extracted_data, extract_errors = extract_data_from_plan(inputs.main_excel, plan)
        errors.extend(extract_errors)
        process_events.append({"step": "extract", "status": "done", "message": f"items={len(extracted_data.get('items', []))}"})
    except Exception as exc:
        errors.append({"category": "extract", "location": "main workbook", "message": str(exc)})
        process_events.append({"step": "extract", "status": "failed", "message": str(exc)})

    try:
        extracted_data, calculation_logs, calculation_errors = apply_calculations(extracted_data)
        errors.extend(calculation_errors)
        process_events.append({"step": "calculate", "status": "done", "message": f"log rows={len(calculation_logs)}"})
    except Exception as exc:
        calculation_logs = []
        errors.append({"category": "calculate", "location": "extracted data", "message": str(exc)})
        process_events.append({"step": "calculate", "status": "failed", "message": str(exc)})

    try:
        if inputs is None:
            raise ValueError("input files are invalid")
        output_files, write_errors = write_filled_templates(
            extracted_data=extracted_data,
            template_files=inputs.templates,
            output_dir=task_paths.filled_files_dir,
            allow_insert_rows=False,
            template_aliases=inputs.template_aliases,
        )
        errors.extend(write_errors)
        process_events.append({"step": "fill", "status": "done", "message": f"filled files={len(output_files)}"})
    except Exception as exc:
        errors.append({"category": "fill", "location": "templates", "message": str(exc)})
        process_events.append({"step": "fill", "status": "failed", "message": str(exc)})

    try:
        if inputs is None:
            raise ValueError("input files are invalid")
        validation_rows, validation_errors = validate_outputs(extracted_data, inputs.templates, output_files)
        errors.extend(validation_errors)
        process_events.append({"step": "validate", "status": "done", "message": f"checks={len(validation_rows)}"})
    except Exception as exc:
        errors.append({"category": "validate", "location": "outputs", "message": str(exc)})
        process_events.append({"step": "validate", "status": "failed", "message": str(exc)})

    try:
        (task_paths.logs_dir / "extracted_data.json").write_text(
            json.dumps(extracted_data, ensure_ascii=False, indent=2, default=str),
            encoding="utf-8",
        )
        write_calculation_log(calculation_logs, task_paths.logs_dir / "calculation_log.xlsx")
        write_process_log(process_events, task_paths.logs_dir / "process_log.xlsx")
        write_error_report(errors, task_paths.reports_dir / "error_report.xlsx")
        write_data_error_report(errors, task_paths.reports_dir / "data_error_report.xlsx")
        write_validation_report(validation_rows, task_paths.reports_dir / "validation_report.xlsx")
        write_summary(
            task_paths.summary_file,
            main_file=inputs.main_excel.name if inputs else "",
            templates=_processed_templates(extracted_data),
            extracted_data=extracted_data,
            output_files=output_files,
            errors=errors,
            validation_rows=validation_rows,
        )
    except Exception as exc:
        print(f"Report writing failed: {exc}")
        return 1

    print(f"Full run completed. Filled files: {len(output_files)}")
    print(f"Validation checks: {len(validation_rows)}")
    print(f"Errors: {len(errors)}")
    print(f"Validation report: {task_paths.reports_dir / 'validation_report.xlsx'}")
    print(f"Error report: {task_paths.reports_dir / 'error_report.xlsx'}")
    print(f"Process log: {task_paths.logs_dir / 'process_log.xlsx'}")
    print(f"Summary: {task_paths.summary_file}")
    return 1 if errors else 0


def _processed_templates(extracted_data: dict[str, object]) -> list[str]:
    templates = []
    for item in extracted_data.get("items", []):
        template = item.get("target_template")
        if template:
            templates.append(str(template))
    return list(dict.fromkeys(templates))


def _build_plan_for_command(config_path: Path, task_paths: object) -> tuple[list[dict[str, object]], list[dict[str, str]], object | None]:
    file_manager = FileManager()
    config, errors = load_mapping_config(config_path)
    inputs = None
    main_structure = None
    template_structure = None

    try:
        inputs = file_manager.collect_inputs()
        errors.extend(_conversion_errors(inputs))
    except Exception as exc:
        errors.append({"category": "input", "location": "input", "message": str(exc)})

    if inputs is not None:
        try:
            main_structure = read_excel_structure(inputs.main_excel)
            template_structure = analyze_templates(inputs.templates)
        except Exception as exc:
            errors.append({"category": "excel", "location": "input files", "message": str(exc)})

    plan, plan_errors = generate_fill_plan(
        config=config,
        main_structure=main_structure,
        template_structure=template_structure,
        input_main_file=inputs.main_excel if inputs else None,
        template_files=inputs.templates if inputs else [],
        template_aliases=inputs.template_aliases if inputs else {},
    )
    errors.extend(plan_errors)

    fill_plan_json = task_paths.logs_dir / "fill_plan.json"
    fill_plan_xlsx = task_paths.logs_dir / "fill_plan.xlsx"
    write_fill_plan_json(plan, fill_plan_json)
    write_fill_plan_xlsx(plan, fill_plan_xlsx)
    return plan, errors, inputs


def _conversion_errors(inputs: object) -> list[dict[str, str]]:
    errors = []
    for record in getattr(inputs, "conversion_records", []):
        if record.status == "failed":
            errors.append(
                {
                    "category": "conversion",
                    "location": record.source_path,
                    "message": f"{record.method}: {record.message}",
                }
            )
    return errors


if __name__ == "__main__":
    raise SystemExit(main())
