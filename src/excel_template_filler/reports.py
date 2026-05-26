from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any


@dataclass
class Issue:
    level: str
    job: str
    rule: str
    message: str
    location: str | None = None


@dataclass
class ValidationItem:
    job: str
    rule: str
    status: str
    message: str
    location: str | None = None


class RunReports:
    def __init__(self, output_root: Path) -> None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.run_dir = output_root / f"run_{timestamp}"
        self.files_dir = self.run_dir / "files"
        self.logs_dir = self.run_dir / "logs"
        self.reports_dir = self.run_dir / "reports"
        self.files_dir.mkdir(parents=True, exist_ok=True)
        self.logs_dir.mkdir(parents=True, exist_ok=True)
        self.reports_dir.mkdir(parents=True, exist_ok=True)
        self.issues: list[Issue] = []
        self.validations: list[ValidationItem] = []
        self.outputs: list[str] = []

    @property
    def log_path(self) -> Path:
        return self.logs_dir / "run.log"

    def add_issue(self, issue: Issue) -> None:
        self.issues.append(issue)

    def add_validation(self, item: ValidationItem) -> None:
        self.validations.append(item)

    def add_output(self, path: Path) -> None:
        self.outputs.append(str(path))

    def write(self) -> None:
        self._write_json("exceptions.json", [asdict(item) for item in self.issues])
        self._write_json("validation.json", [asdict(item) for item in self.validations])
        self._write_json(
            "summary.json",
            {
                "status": "failed" if self.issues else "success",
                "outputs": self.outputs,
                "issue_count": len(self.issues),
                "validation_count": len(self.validations),
            },
        )

    def _write_json(self, filename: str, payload: Any) -> None:
        path = self.reports_dir / filename
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
