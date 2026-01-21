"""Listado de apps únicas para archivos Office con permiso de copia."""
from __future__ import annotations

import argparse
import os
import subprocess
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


def launch_apps(apps: list[str], design_mode: bool) -> None:
    if not design_mode:
        return
    if os.name != "nt":
        print("[WARN] Apertura de aplicaciones omitida: no es Windows.")
        return
    mapping = {
        "WORD": "winword.exe",
        "POWERPOINT": "powerpnt.exe",
        "EXCEL": "excel.exe",
    }
    for app in apps:
        exe = mapping.get(app)
        if not exe:
            continue
        try:
            print(f"[OPEN] Intentando abrir {app} ({exe}) con startfile.")
            os.startfile(exe)  # type: ignore[arg-type]
        except OSError as exc:
            print(f"[WARN] No se pudo iniciar {app} con startfile ({exc}); reintentando con cmd.")
            try:
                print(f"[OPEN] Intentando abrir {app} ({exe}) con cmd start.")
                subprocess.run(["cmd", "/c", "start", "", exe], check=False)
            except OSError as retry_exc:
                print(f"[WARN] No se pudo iniciar {app} con cmd ({retry_exc})")


def run_actions(base_dir: Path, design_mode: bool) -> list[str]:
    apps = iter_copy_allowed_apps(base_dir)
    launch_apps(apps, design_mode)
    if design_mode:
        print({"apps": apps})
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
    parser.add_argument(
        "--design-mode",
        action="store_true",
        help="Muestra información de depuración y abre aplicaciones.",
    )
    args = parser.parse_args(argv)
    base_dir = path_utils.normalize_path(Path(args.base_dir)).resolve()
    design_mode = args.design_mode
    run_actions(base_dir, design_mode)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
