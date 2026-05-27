"""File discovery and task directory utilities."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from src.xls_converter import ConversionRecord, convert_xls_file, write_conversion_log

SUPPORTED_EXCEL_EXTENSIONS = {".xlsx", ".xlsm"}
CONVERTIBLE_EXCEL_EXTENSIONS = {".xls"}


@dataclass(frozen=True)
class TaskPaths:
    root: Path
    filled_files_dir: Path
    logs_dir: Path
    reports_dir: Path
    summary_file: Path
    task_dir: Path
    backups_dir: Path
    converted_dir: Path


@dataclass(frozen=True)
class InputFiles:
    main_excel: Path
    templates: list[Path]
    converted_files: list[Path]
    conversion_records: list[ConversionRecord]
    template_aliases: dict[str, Path]


class FileManager:
    def __init__(
        self,
        main_excel_dir: str | Path = "input/main_excel",
        template_dir: str | Path = "input/templates",
        output_dir: str | Path = "output",
    ) -> None:
        self.main_excel_dir = Path(main_excel_dir)
        self.template_dir = Path(template_dir)
        self.output_dir = Path(output_dir)

    def create_task_dirs(self) -> TaskPaths:
        filled_files_dir = self.output_dir / "filled_files"
        logs_dir = self.output_dir / "logs"
        reports_dir = self.output_dir / "reports"
        summary_file = self.output_dir / "summary.txt"
        task_dir = self.output_dir / "tasks" / datetime.now().strftime("%Y%m%d_%H%M%S")
        backups_dir = self.output_dir / "backups"
        converted_dir = self.output_dir / "converted"

        for path in (filled_files_dir, logs_dir, reports_dir, task_dir, backups_dir, converted_dir):
            path.mkdir(parents=True, exist_ok=True)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        if not summary_file.exists():
            summary_file.write_text("", encoding="utf-8")

        return TaskPaths(
            root=self.output_dir,
            filled_files_dir=filled_files_dir,
            logs_dir=logs_dir,
            reports_dir=reports_dir,
            summary_file=summary_file,
            task_dir=task_dir,
            backups_dir=backups_dir,
            converted_dir=converted_dir,
        )

    def collect_inputs(self) -> InputFiles:
        task_paths = self.create_task_dirs()
        main_files, main_xls = self._collect_excel_files(self.main_excel_dir)
        template_files, template_xls = self._collect_excel_files(self.template_dir)
        conversion_records: list[ConversionRecord] = []
        converted_files: list[Path] = []
        template_aliases: dict[str, Path] = {}

        for xls_path in main_xls + template_xls:
            record = convert_xls_file(xls_path, task_paths.converted_dir, task_paths.backups_dir)
            conversion_records.append(record)
            if record.status == "converted":
                converted_path = Path(record.output_path)
                converted_files.append(converted_path)
                if xls_path.parent == self.main_excel_dir:
                    main_files.append(converted_path)
                else:
                    template_files.append(converted_path)
                    template_aliases[xls_path.name] = converted_path

        write_conversion_log(conversion_records, task_paths.logs_dir / "conversion_log.xlsx")

        if len(main_files) != 1:
            failed_main = [record for record in conversion_records if Path(record.source_path).parent == self.main_excel_dir and record.status == "failed"]
            if failed_main:
                messages = "; ".join(record.message for record in failed_main)
                raise ValueError(f"main .xls conversion failed: {messages}")
            raise ValueError(
                f"input/main_excel must contain exactly one supported main Excel file; found {len(main_files)}"
            )
        if not template_files:
            raise ValueError("input/templates must contain at least one supported template Excel file")

        return InputFiles(
            main_excel=main_files[0],
            templates=template_files,
            converted_files=converted_files,
            conversion_records=conversion_records,
            template_aliases=template_aliases,
        )

    def _collect_excel_files(self, directory: Path) -> tuple[list[Path], list[Path]]:
        if not directory.exists():
            raise FileNotFoundError(f"Directory does not exist: {directory}")

        supported: list[Path] = []
        convertible: list[Path] = []
        for path in sorted(directory.iterdir()):
            if not path.is_file() or path.name.startswith("~$"):
                continue
            suffix = path.suffix.lower()
            if suffix in SUPPORTED_EXCEL_EXTENSIONS:
                supported.append(path)
            elif suffix in CONVERTIBLE_EXCEL_EXTENSIONS:
                convertible.append(path)
        return supported, convertible
