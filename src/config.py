from __future__ import annotations

from dataclasses import dataclass
import os
from pathlib import Path

from dotenv import load_dotenv
import yaml


BASE_DIR = Path(__file__).resolve().parents[1]
load_dotenv(BASE_DIR / ".env")


@dataclass(frozen=True)
class Settings:
    base_dir: Path = BASE_DIR
    gemini_api_key: str = os.getenv("GEMINI_API_KEY", "").strip()
    gemini_model: str = os.getenv("GEMINI_MODEL", "gemini-3.8-flash").strip()
    save_uploads: bool = os.getenv("SAVE_UPLOADS", "false").strip().lower() in {"1", "true", "yes"}

    @property
    def templates_dir(self) -> Path:
        return self.base_dir / "Templates"

    @property
    def output_dir(self) -> Path:
        return self.base_dir / "Output"

    @property
    def uploads_dir(self) -> Path:
        return self.base_dir / "data_uploads"

    @property
    def mapping_path(self) -> Path:
        return self.base_dir / "config" / "excel_mapping.yaml"


def load_mapping(path: Path | None = None) -> dict:
    mapping_path = path or Settings().mapping_path
    if not mapping_path.exists():
        return {}
    with mapping_path.open("r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def discover_workbook(mapping: dict | None = None, base_dir: Path | None = None) -> Path:
    """Resolve the configured workbook, or the only supported template."""
    mapping = mapping if mapping is not None else load_mapping()
    root = Path(base_dir or BASE_DIR)
    configured = mapping.get("workbook")
    if configured:
        path = (root / configured).resolve()
        if path.is_file():
            return path

    candidates = sorted(
        p for p in (root / "Templates").glob("*")
        if p.suffix.lower() in {".xlsx", ".xlsm"} and not p.name.startswith("~$")
    )
    if not candidates:
        if configured:
            raise FileNotFoundError(f"Configured workbook was not found and no .xlsx or .xlsm workbook is available in Templates: {root / 'Templates'}")
        raise FileNotFoundError("No .xlsx or .xlsm workbook was found in Templates.")
    if len(candidates) > 1:
        raise RuntimeError("More than one Excel workbook is in Templates and the configured path is missing. Select one in config/excel_mapping.yaml.")
    return candidates[0]
