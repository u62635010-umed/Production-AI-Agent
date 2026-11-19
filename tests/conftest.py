from pathlib import Path

import pytest
from openpyxl import Workbook

from src.config import Settings
from src.excel_manager import ExcelManager


@pytest.fixture
def workbook_env(tmp_path: Path):
    templates = tmp_path / "Templates"
    templates.mkdir()
    output = tmp_path / "Output"
    output.mkdir()
    path = templates / "production_sheet.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Quality"
    ws["A1"] = "=1+1"
    ws["B1"] = "preserve this"
    ws["B1"].number_format = "@"
    wb.create_sheet("production_summary")
    wb.save(path)
    settings = Settings(base_dir=tmp_path)
    return path, settings, ExcelManager(path, settings)
