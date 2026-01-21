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


def open_destinations(destinations: list[str], design_mode: bool) -> None:
    if os.name != "nt":
        if design_mode:
            print("[WARN] Apertura de carpetas omitida: no es Windows.")
        return
    for destination in destinations:
        try:
            if design_mode:
                print(f"[OPEN] Intentando abrir carpeta {destination} con startfile.")
            os.startfile(destination)  # type: ignore[arg-type]
        except OSError as exc:
            if design_mode:
                print(
                    f"[WARN] No se pudo abrir carpeta con startfile ({exc}); reintentando con cmd."
                )
            try:
                if design_mode:
                    print(f"[OPEN] Intentando abrir carpeta {destination} con cmd start.")
                subprocess.run(["cmd", "/c", "start", "", destination], check=False)
            except OSError as retry_exc:
                if design_mode:
                    print(f"[WARN] No se pudo abrir carpeta con cmd ({retry_exc})")


def run_actions(base_dir: Path, design_mode: bool) -> list[str]:
    destinations = iter_copy_allowed_destinations(base_dir)
    open_destinations(destinations, design_mode)
    if design_mode:
        print({"destinations": destinations})
    return destinations


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
    parser.add_argument(
        "--design-mode",
        action="store_true",
        help="Muestra información de depuración y abre destinos.",
    )
    args = parser.parse_args(argv)
    base_dir = path_utils.normalize_path(Path(args.base_dir)).resolve()
    design_mode = args.design_mode
    run_actions(base_dir, design_mode)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
