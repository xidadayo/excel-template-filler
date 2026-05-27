"""Parse and validate mapping configuration files."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, Field, ValidationError


class MappingRule(BaseModel):
    source_field: str = Field(min_length=1)
    target_cell: str = Field(min_length=1)
    write_type: Literal["single_value"]
    required: bool = True
    overwrite_formula: bool = False


class CalculationRule(BaseModel):
    target_field: str = Field(min_length=1)
    formula: str = Field(min_length=1)


class DetailAreaConfig(BaseModel):
    source_sheet: str = Field(min_length=1)
    target_sheet: str = Field(min_length=1)
    target_start_row: int = Field(ge=1)
    target_end_row: int = Field(ge=1)
    columns: dict[str, str] = Field(default_factory=dict)
    calculations: list[CalculationRule] = Field(default_factory=list)


class TemplateMappingConfig(BaseModel):
    template_name: str = Field(min_length=1)
    output_name: str = Field(min_length=1)
    source_sheet: str = Field(min_length=1)
    target_sheet: str = Field(min_length=1)
    mappings: list[MappingRule] = Field(default_factory=list)
    detail_area: DetailAreaConfig | None = None


class MappingConfig(BaseModel):
    project_name: str = Field(min_length=1)
    main_file: str = Field(min_length=1)
    templates: list[TemplateMappingConfig] = Field(default_factory=list)


def load_mapping_config(path: str | Path) -> tuple[MappingConfig | None, list[dict[str, str]]]:
    config_path = Path(path)
    errors: list[dict[str, str]] = []

    try:
        raw = json.loads(config_path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return None, [_error("config", str(config_path), "mapping config file not found")]
    except json.JSONDecodeError as exc:
        return None, [_error("config", str(config_path), f"invalid JSON: {exc}")]

    try:
        return MappingConfig.model_validate(raw), errors
    except ValidationError as exc:
        for item in exc.errors():
            location = ".".join(str(part) for part in item["loc"])
            errors.append(_error("config", location, item["msg"]))
        return None, errors


def _error(category: str, location: str, message: str) -> dict[str, str]:
    return {"category": category, "location": location, "message": message}
