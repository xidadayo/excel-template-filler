"""Convert legacy .xls workbooks into .xlsx files when possible."""

from __future__ import annotations

import os
import shutil
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

from openpyxl import Workbook


ConversionStatus = Literal["converted", "failed", "skipped"]


@dataclass(frozen=True)
class ConversionRecord:
    source_path: str
    backup_path: str
    output_path: str
    method: str
    status: ConversionStatus
    message: str


def convert_xls_file(source_path: Path, output_dir: Path, backup_dir: Path) -> ConversionRecord:
    output_dir.mkdir(parents=True, exist_ok=True)
    backup_dir.mkdir(parents=True, exist_ok=True)

    backup_path = backup_dir / source_path.name
    shutil.copy2(source_path, backup_path)
    output_path = output_dir / f"{source_path.stem}.xlsx"

    method = detect_conversion_method()
    if method == "none":
        return ConversionRecord(
            source_path=str(source_path),
            backup_path=str(backup_path),
            output_path=str(output_path),
            method=method,
            status="failed",
            message="No supported converter found. Install Microsoft Excel with pywin32/xlwings or LibreOffice.",
        )

    try:
        if method == "excel_com":
            _convert_with_excel_com(source_path, output_path)
        elif method == "libreoffice":
            _convert_with_libreoffice(source_path, output_dir)
            converted_path = output_dir / f"{source_path.stem}.xlsx"
            if converted_path != output_path and converted_path.exists():
                converted_path.replace(output_path)
        else:
            raise RuntimeError(f"Unsupported conversion method: {method}")
    except Exception as exc:
        return ConversionRecord(
            source_path=str(source_path),
            backup_path=str(backup_path),
            output_path=str(output_path),
            method=method,
            status="failed",
            message=str(exc),
        )

    if not output_path.exists():
        return ConversionRecord(
            source_path=str(source_path),
            backup_path=str(backup_path),
            output_path=str(output_path),
            method=method,
            status="failed",
            message="Converter finished but output .xlsx was not created.",
        )

    return ConversionRecord(
        source_path=str(source_path),
        backup_path=str(backup_path),
        output_path=str(output_path),
        method=method,
        status="converted",
        message="Converted successfully.",
    )


def detect_conversion_method() -> str:
    if os.name == "nt" and _has_excel_com():
        return "excel_com"
    if _find_libreoffice() is not None:
        return "libreoffice"
    return "none"


def write_conversion_log(records: list[ConversionRecord], path: str | Path) -> None:
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "conversion_log"
    sheet.append(["source_path", "backup_path", "output_path", "method", "status", "message"])
    for record in records:
        sheet.append(
            [
                record.source_path,
                record.backup_path,
                record.output_path,
                record.method,
                record.status,
                record.message,
            ]
        )
    workbook.save(output_path)


def _has_excel_com() -> bool:
    try:
        import win32com.client  # type: ignore  # noqa: F401
    except Exception:
        return False
    return True


def _find_libreoffice() -> str | None:
    for command in ("soffice", "libreoffice"):
        found = shutil.which(command)
        if found:
            return found
    common_windows_paths = [
        Path("C:/Program Files/LibreOffice/program/soffice.exe"),
        Path("C:/Program Files (x86)/LibreOffice/program/soffice.exe"),
    ]
    for path in common_windows_paths:
        if path.exists():
            return str(path)
    return None


def _convert_with_excel_com(source_path: Path, output_path: Path) -> None:
    import win32com.client  # type: ignore

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.DisplayAlerts = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(source_path.resolve()))
        workbook.SaveAs(str(output_path.resolve()), FileFormat=51)
    finally:
        if workbook is not None:
            workbook.Close(False)
        excel.Quit()


def _convert_with_libreoffice(source_path: Path, output_dir: Path) -> None:
    executable = _find_libreoffice()
    if executable is None:
        raise RuntimeError("LibreOffice executable not found.")
    result = subprocess.run(
        [
            executable,
            "--headless",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(output_dir.resolve()),
            str(source_path.resolve()),
        ],
        capture_output=True,
        text=True,
        timeout=120,
        check=False,
    )
    if result.returncode != 0:
        raise RuntimeError((result.stderr or result.stdout or "LibreOffice conversion failed.").strip())
