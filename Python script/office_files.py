"""Listado inicial de archivos Office en una carpeta."""
from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable

OFFICE_EXTENSIONS = (
    ".dotx",
    ".dotm",
    ".potx",
    ".potm",
    ".xltx",
    ".xltm",
    ".thmx",
)


def iter_office_files(base_dir: Path, extensions: Iterable[str] = OFFICE_EXTENSIONS) -> list[dict[str, str]]:
    base_dir = Path(base_dir)
    items: list[dict[str, str]] = []
    for ext in extensions:
        for path in base_dir.glob(f"*{ext}"):
            if not path.is_file():
                continue
            items.append(
                {
                    "name": path.name,
                    "path": str(path),
                    "extension": path.suffix.lower(),
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
    base_dir = Path(args.base_dir)
    items = iter_office_files(base_dir)
    for item in items:
        print(f"{item['extension']}\t{item['name']}\t{item['path']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
