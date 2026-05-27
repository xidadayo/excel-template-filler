# User Guide

## 1. Prepare Files

Put exactly one main Excel file in:

```text
input/main_excel/
```

Put one or more template files in:

```text
input/templates/
```

Supported formats:

- `.xlsx`
- `.xlsm`
- `.xls` through conversion

## 2. Prepare Config

Copy:

```text
config/mapping_config.example.json
```

to:

```text
config/mapping_config.json
```

Update file names, Sheet names, fields, target cells, and detail area settings.

## 3. Analyze Workbooks

```powershell
python main.py analyze
```

Check:

- `output/logs/main_excel_structure.json`
- `output/logs/template_structure.json`

## 4. Generate Fill Plan

```powershell
python main.py plan --config config/mapping_config.json
```

Check:

- `output/logs/fill_plan.xlsx`
- `output/reports/config_error_report.xlsx`

## 5. Extract Data

```powershell
python main.py extract --config config/mapping_config.json
```

Check:

- `output/logs/extracted_data.json`
- `output/logs/calculation_log.xlsx`
- `output/reports/data_error_report.xlsx`

## 6. Fill Templates

```powershell
python main.py fill --config config/mapping_config.json
```

Check:

- `output/filled_files/`
- `output/reports/error_report.xlsx`
- `output/summary.txt`

## 7. Run Everything

```powershell
python main.py run --config config/mapping_config.json
```

This is the recommended production command.

## 8. Natural-Language Rule Draft

```powershell
python main.py parse-rules --rules config/sample_rules_prompt.txt
```

Review:

- `config/mapping_config.draft.json`
- `config/mapping_config.need_confirm.json`

Only use the draft after manual confirmation.

## 9. Web Page

```powershell
python -m uvicorn web_app:app --host 127.0.0.1 --port 8000
```

Open:

```text
http://127.0.0.1:8000/
```

Use the page to upload the main Excel file and template files, then type natural-language rules directly into the form. The web layer generates `config/mapping_config.json` and `config/mapping_config.need_confirm.json`, then you can run analysis, generate plans, execute filling, and download results.

## 10. Reports

- `summary.txt`: quick run summary.
- `error_report.xlsx`: all errors and warnings.
- `validation_report.xlsx`: output validation checks.
- `process_log.xlsx`: workflow steps.
- `conversion_log.xlsx`: `.xls` conversion details.
