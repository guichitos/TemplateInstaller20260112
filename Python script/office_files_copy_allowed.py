"""Listado de archivos Office con permiso de copia."""
from __future__ import annotations

import argparse
from pathlib import Path

import office_files
import path_utils


def iter_copy_allowed_files(base_dir: Path) -> list[dict[str, str]]:
    items = office_files.iter_office_files(base_dir)
    return [item for item in items if item.get("copy") == "true"]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Listado de archivos Office con permiso de copia (copy=true).",
    )
    parser.add_argument(
        "base_dir",
        nargs="?",
        default=".",
        help="Carpeta a escanear (por defecto, la carpeta actual).",
    )
    args = parser.parse_args(argv)
    base_dir = path_utils.normalize_path(Path(args.base_dir)).resolve()
    items = iter_copy_allowed_files(base_dir)
    name_width = max((len(item["name"]) for item in items), default=len("name"))
    print(f"{'name':<{name_width}}  {'extension':<9}  {'copy':<5}  {'app':<10}  destination")
    for item in items:
        print(
            f"{item['name']:<{name_width}}  "
            f"{item['extension']:<9}  "
            f"{item['copy']:<5}  "
            f"{item['app']:<10}  "
            f"{item['destination']}"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
