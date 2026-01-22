"""List unique apps for Office files that allow copying."""
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
    if os.name != "nt":
        if design_mode:
            print("[WARN] Skipping app launch: not on Windows.")
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
            if design_mode:
                print(f"[OPEN] Attempting to open {app} ({exe}) with startfile.")
            os.startfile(exe)  # type: ignore[arg-type]
        except OSError as exc:
            if design_mode:
                print(f"[WARN] Could not start {app} with startfile ({exc}); retrying with cmd.")
            try:
                if design_mode:
                    print(f"[OPEN] Attempting to open {app} ({exe}) with cmd start.")
                subprocess.run(["cmd", "/c", "start", "", exe], check=False)
            except OSError as retry_exc:
                if design_mode:
                    print(f"[WARN] Could not start {app} with cmd ({retry_exc})")


def run_actions(base_dir: Path, design_mode: bool) -> list[str]:
    apps = iter_copy_allowed_apps(base_dir)
    launch_apps(apps, design_mode)
    if design_mode:
        print({"apps": apps})
    return apps


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="List unique apps for Office files that allow copying.",
    )
    parser.add_argument(
        "base_dir",
        nargs="?",
        default=".",
        help="Folder to scan (defaults to the current folder).",
    )
    parser.add_argument(
        "--design-mode",
        action="store_true",
        help="Show debug information and open apps.",
    )
    args = parser.parse_args(argv)
    base_dir = path_utils.normalize_path(Path(args.base_dir)).resolve()
    design_mode = args.design_mode
    run_actions(base_dir, design_mode)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
