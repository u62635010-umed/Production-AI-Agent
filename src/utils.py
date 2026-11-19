from __future__ import annotations

from hashlib import sha256
import re
from pathlib import Path


def image_sha256(data: bytes) -> str:
    return sha256(data).hexdigest()


def safe_filename(name: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_.-]+", "_", Path(name).name).strip("._")
    return cleaned[:120] or "production_photo.png"
