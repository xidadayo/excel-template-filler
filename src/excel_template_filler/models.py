from __future__ import annotations

from pathlib import Path
from typing import Literal

from pydantic import BaseModel, Field, field_validator


class LookupSource(BaseModel):
    kind: Literal["lookup"] = "lookup"
    sheet: str
    key_column: str
    key_value: str
    value_column: str
    header_row: int = Field(default=1, ge=1)


class CellSource(BaseModel):
    kind: Literal["cell"] = "cell"
    sheet: str
    cell: str


class TableSource(BaseModel):
    kind: Literal["table"] = "table"
    sheet: str
    columns: list[str]
    header_row: int = Field(default=1, ge=1)
    start_row: int = Field(default=2, ge=1)


class CellTarget(BaseModel):
    sheet: str
    cell: str


class RangeTarget(BaseModel):
    sheet: str
    start_cell: str
    max_rows: int = Field(ge=1)


class CellWriteRule(BaseModel):
    type: Literal["cell"] = "cell"
    name: str
    source: LookupSource | CellSource
    target: CellTarget
    allow_overwrite_formula: bool = False


class RangeWriteRule(BaseModel):
    type: Literal["range"] = "range"
    name: str
    source: TableSource
    target: RangeTarget
    allow_overwrite_formula: bool = False


WriteRule = CellWriteRule | RangeWriteRule


class TemplateJob(BaseModel):
    name: str
    template_path: Path
    output_name: str
    writes: list[WriteRule]

    @field_validator("output_name")
    @classmethod
    def output_must_be_xlsx(cls, value: str) -> str:
        if not value.lower().endswith(".xlsx"):
            raise ValueError("output_name must end with .xlsx")
        return value


class AppConfig(BaseModel):
    master_path: Path
    output_dir: Path = Path("outputs")
    jobs: list[TemplateJob]

    @field_validator("master_path")
    @classmethod
    def master_must_exist(cls, value: Path) -> Path:
        if not value.exists():
            raise ValueError(f"master_path does not exist: {value}")
        return value

    @field_validator("jobs")
    @classmethod
    def jobs_must_not_be_empty(cls, value: list[TemplateJob]) -> list[TemplateJob]:
        if not value:
            raise ValueError("jobs cannot be empty")
        return value
