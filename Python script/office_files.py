"""Listado inicial de archivos Office en una carpeta."""
from __future__ import annotations

import argparse
import os
from pathlib import Path
from typing import Iterable

import path_utils

OFFICE_EXTENSIONS = (
    ".dotx",
    ".dotm",
    ".potx",
    ".potm",
    ".xltx",
    ".xltm",
    ".thmx",
)

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


def _resolve_template_paths() -> dict[str, Path]:
    base_paths = path_utils.resolve_base_paths()
    custom_word = path_utils.normalize_path(
        os.environ.get("CUSTOM_OFFICE_TEMPLATE_PATH", base_paths["CUSTOM_WORD"])
    )
    custom_ppt = path_utils.normalize_path(
        os.environ.get("POWERPOINT_TEMPLATE_PATH", base_paths["CUSTOM_PPT"])
    )
    custom_excel = path_utils.normalize_path(
        os.environ.get("EXCEL_TEMPLATE_PATH", base_paths["CUSTOM_EXCEL"])
    )
    roaming = path_utils.normalize_path(
        os.environ.get("ROAMING_TEMPLATE_FOLDER_PATH", base_paths["ROAMING"])
    )
    excel_startup = path_utils.normalize_path(
        os.environ.get("EXCEL_STARTUP_FOLDER_PATH", base_paths["EXCEL_STARTUP"])
    )
    theme = path_utils.normalize_path(base_paths["THEME"])
    return {
        "THEME": theme,
        "CUSTOM_WORD": custom_word,
        "CUSTOM_PPT": custom_ppt or custom_word,
        "CUSTOM_EXCEL": custom_excel or custom_word,
        "ROAMING": roaming,
        "EXCEL": excel_startup,
    }


def _destination_for_file(name: str, extension: str, paths: dict[str, Path]) -> Path | None:
    if name in BASE_TEMPLATE_NAMES:
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


def iter_office_files(base_dir: Path, extensions: Iterable[str] = OFFICE_EXTENSIONS) -> list[dict[str, str]]:
    base_dir = Path(base_dir)
    paths = _resolve_template_paths()
    items: list[dict[str, str]] = []
    for ext in extensions:
        for path in base_dir.glob(f"*{ext}"):
            if not path.is_file():
                continue
            destination_root = _destination_for_file(path.name, path.suffix.lower(), paths)
            items.append(
                {
                    "name": path.name,
                    "path": str(path),
                    "extension": path.suffix.lower(),
                    "destination": str((destination_root / path.name).resolve()) if destination_root else "",
                }
            )
    return items


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Listado de archivos Office en una carpeta.")
    parser.add_argument(
        "base_dir",
        nargs="?",
        default=".",
        help="Carpeta a escanear (por defecto, la carpeta actual).",
    )
    args = parser.parse_args(argv)
    base_dir = path_utils.normalize_path(Path(args.base_dir)).resolve()
    items = iter_office_files(base_dir)
    name_width = max((len(item["name"]) for item in items), default=len("name"))
    print(f"{'name':<{name_width}}  {'extension':<9}  destination")
    for item in items:
        print(f"{item['name']:<{name_width}}  {item['extension']:<9}  {item['destination']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
