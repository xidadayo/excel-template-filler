# Project Design

## Overview

`excel-template-filler` is a local Excel automation system. It reads one main workbook and multiple template workbooks, generates a fill plan from JSON configuration, extracts and cleans source data, performs simple calculations, safely writes values into copied templates, and produces reports.

The system is intentionally conservative: original Excel files are never modified, and template copies are written only through `cell.value`.

## Architecture

```text
main.py / web_app.py
        |
        v
src/file_manager.py
src/xls_converter.py
        |
        v
src/excel_reader.py
src/template_analyzer.py
        |
        v
src/rule_parser.py
src/fill_plan_generator.py
        |
        v
src/data_extractor.py
src/data_cleaner.py
src/data_calculator.py
        |
        v
src/excel_writer.py
src/validator.py
        |
        v
src/report_writer.py
src/exception_handler.py
```

## Module Responsibilities

- `file_manager.py`: discovers input files, creates output folders, coordinates `.xls` conversion.
- `xls_converter.py`: backs up and converts `.xls` files using Excel COM or LibreOffice.
- `excel_reader.py`: reads main workbook structure without modifying files.
- `template_analyzer.py`: analyzes template sheets, merged cells, formulas, non-empty cells, and possible detail areas.
- `rule_parser.py`: parses and validates `mapping_config.json`.
- `nl_rule_parser.py`: parses natural-language rules into a draft mapping config.
- `fill_plan_generator.py`: generates `fill_plan.json` and `fill_plan.xlsx`.
- `data_extractor.py`: extracts single values and detail columns from the main workbook.
- `data_cleaner.py`: normalizes text, numbers, dates, and empty values while preserving originals.
- `data_calculator.py`: evaluates simple formulas and writes calculation logs.
- `excel_writer.py`: copies templates and writes only target `cell.value`.
- `validator.py`: validates outputs against original templates and extracted data.
- `report_writer.py`: writes validation, error, process, and summary reports.
- `web_app.py`: provides a local FastAPI page that calls existing workflow functions.

## Workflow

1. Analyze input files.
2. Parse configuration.
3. Generate fill plan.
4. Extract and clean data.
5. Calculate derived fields.
6. Copy templates.
7. Write values into copied templates.
8. Validate outputs.
9. Write logs, reports, and summary.

## Output Files

- `output/logs/main_excel_structure.json`
- `output/logs/template_structure.json`
- `output/logs/fill_plan.json`
- `output/logs/fill_plan.xlsx`
- `output/logs/extracted_data.json`
- `output/logs/calculation_log.xlsx`
- `output/logs/process_log.xlsx`
- `output/logs/conversion_log.xlsx`
- `output/reports/config_error_report.xlsx`
- `output/reports/data_error_report.xlsx`
- `output/reports/error_report.xlsx`
- `output/reports/validation_report.xlsx`
- `output/summary.txt`
- `output/filled_files/*.xlsx`

## Safety Guarantees

- Original input files are read-only.
- Templates are copied before writing.
- Only configured target cells/ranges are written.
- Only `cell.value` is changed.
- Formula cells are not overwritten unless explicitly allowed.
- Rows are not inserted by default.
- Missing fields are reported, not guessed.
