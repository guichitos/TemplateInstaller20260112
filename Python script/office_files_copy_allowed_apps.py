"""Listado de apps únicas para archivos Office con permiso de copia."""
from __future__ import annotations

import argparse
from pathlib import Path

import office_files_copy_allowed
import path_utils


def iter_copy_allowed_apps(base_dir: Path) -> list[str]:
    items = office_files_copy_allowed.iter_copy_allowed_files(base_dir)
    seen: set[str] = set()
    apps: list[str] = []
    for item in items:
        app = item.get("app", "")
        if not app or app in seen:
            continue
        seen.add(app)
        apps.append(app)
    return apps


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Listado único de apps para archivos Office con permiso de copia.",
    )
    parser.add_argument(
        "base_dir",
        nargs="?",
        default=".",
        help="Carpeta a escanear (por defecto, la carpeta actual).",
    )
    args = parser.parse_args(argv)
    base_dir = path_utils.normalize_path(Path(args.base_dir)).resolve()
    apps = iter_copy_allowed_apps(base_dir)
    print({"apps": apps})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
