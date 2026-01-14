"""Listado de destinos únicos para archivos Office con permiso de copia."""
from __future__ import annotations

import argparse
import os
import subprocess
from pathlib import Path

import office_files_copy_allowed
import path_utils


def iter_copy_allowed_destinations(base_dir: Path) -> list[str]:
    items = office_files_copy_allowed.iter_copy_allowed_files(base_dir)
    seen: set[str] = set()
    destinations: list[str] = []
    for item in items:
        destination = item.get("destination", "")
        if not destination or destination in seen:
            continue
        seen.add(destination)
        destinations.append(destination)
    return destinations


def open_destinations(destinations: list[str]) -> None:
    if os.name != "nt":
        print("[WARN] Apertura de carpetas omitida: no es Windows.")
        return
    for destination in destinations:
        try:
            print(f"[OPEN] Intentando abrir carpeta {destination} con startfile.")
            os.startfile(destination)  # type: ignore[arg-type]
        except OSError as exc:
            print(
                f"[WARN] No se pudo abrir carpeta con startfile ({exc}); reintentando con cmd."
            )
            try:
                print(f"[OPEN] Intentando abrir carpeta {destination} con cmd start.")
                subprocess.run(["cmd", "/c", "start", "", destination], check=False)
            except OSError as retry_exc:
                print(f"[WARN] No se pudo abrir carpeta con cmd ({retry_exc})")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Listado único de destinos para archivos Office con permiso de copia.",
    )
    parser.add_argument(
        "base_dir",
        nargs="?",
        default=".",
        help="Carpeta a escanear (por defecto, la carpeta actual).",
    )
    args = parser.parse_args(argv)
    base_dir = path_utils.normalize_path(Path(args.base_dir)).resolve()
    destinations = iter_copy_allowed_destinations(base_dir)
    open_destinations(destinations)
    print({"destinations": destinations})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
