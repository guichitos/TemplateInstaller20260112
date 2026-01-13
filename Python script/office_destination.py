"""Resolución del destino de plantillas Office."""
from __future__ import annotations

from pathlib import Path
from typing import Iterable

BASE_TEMPLATE_NAMES = {
    "Normal.dotx",
    "Normal.dotm",
    "NormalEmail.dotx",
    "NormalEmail.dotm",
    "Blank.potx",
    "Blank.potm",
    "Book.xltx",
    "Book.xltm",
    "Sheet.xltx",
    "Sheet.xltm",
}


def resolve_destination_for_name(
    name: str,
    paths: dict[str, Path],
    base_names: Iterable[str] = BASE_TEMPLATE_NAMES,
) -> Path | None:
    extension = Path(name).suffix.lower()
    if name in base_names:
        if name.startswith(("Normal.", "NormalEmail.", "Blank.")):
            return paths["ROAMING"]
        if name.startswith(("Book.", "Sheet.")):
            return paths["EXCEL"]
    if extension in {".dotx", ".dotm"}:
        return paths["CUSTOM_WORD"]
    if extension in {".potx", ".potm"}:
        return paths["CUSTOM_PPT"]
    if extension in {".xltx", ".xltm"}:
        return paths["CUSTOM_EXCEL"]
    if extension == ".thmx":
        return paths["THEME"]
    return None
