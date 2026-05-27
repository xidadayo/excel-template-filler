# Developer Guide

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
pytest
```

## Commands

```powershell
python main.py analyze
python main.py plan --config config/mapping_config.json
python main.py extract --config config/mapping_config.json
python main.py fill --config config/mapping_config.json
python main.py run --config config/mapping_config.json
python main.py parse-rules --rules config/sample_rules_prompt.txt
python -m uvicorn web_app:app --host 127.0.0.1 --port 8000
```

## Development Rules

- Keep Excel business logic in `src/`.
- Keep `web_app.py` as a thin wrapper over existing workflow functions. It may call `nl_rule_parser` to turn natural-language web input into a draft config, but it must not reimplement Excel processing.
- Never modify original templates.
- Write only to copied files in `output/filled_files/`.
- Write only `cell.value`.
- Do not guess missing fields.
- Add tests for every safety rule.

## Testing Strategy

Tests cover:

- File discovery and `.xls` conversion handling.
- Workbook structure analysis.
- Fill plan generation.
- Data extraction, cleaning, and calculation.
- Formula and template protection.
- Validation reports.
- Natural-language rule parsing.
- FastAPI upload and page rendering.

Run all tests:

```powershell
pytest
```

## Key Extension Points

- Add more calculation functions in `src/data_calculator.py`.
- Add stricter template diff checks in `src/validator.py`.
- Add richer natural-language parsing in `src/nl_rule_parser.py`.
- Add job isolation and multi-user task folders in `web_app.py`.
