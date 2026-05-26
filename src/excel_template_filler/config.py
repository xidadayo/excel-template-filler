from __future__ import annotations

import json
from pathlib import Path

from .models import AppConfig


def load_config(path: Path) -> AppConfig:
    data = json.loads(path.read_text(encoding="utf-8"))
    base_dir = path.parent.resolve()
    data["master_path"] = _resolve(base_dir, data["master_path"])
    data["output_dir"] = _resolve(base_dir, data.get("output_dir", "outputs"))
    for job in data.get("jobs", []):
        job["template_path"] = _resolve(base_dir, job["template_path"])
    return AppConfig.model_validate(data)


def _resolve(base_dir: Path, value: str) -> str:
    path = Path(value)
    if path.is_absolute():
        return str(path)
    return str((base_dir / path).resolve())
