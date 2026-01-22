"""Single uninstaller based on the current folder."""
from __future__ import annotations

import argparse
import logging
from pathlib import Path

# Manual configuration for design mode.
# - Set to True to force design mode on.
# - Set to False to force design mode off.
# - Leave as None to use the normal environment-based logic.
MANUAL_IS_DESIGN_MODE: bool | None = None

try:
    from . import common
except ImportError:  # pragma: no cover - allow direct execution as a script
    import sys

    sys.path.append(str(Path(__file__).resolve().parent))
    import common  # type: ignore[no-redef]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Office template uninstaller (Python)")
    return parser.parse_args()


def main(argv: list[str] | None = None) -> int:
    args = parse_args()
    design_mode = _resolve_design_mode()
    common.refresh_design_log_flags(design_mode)
    common.configure_logging(design_mode)
    common.close_office_apps(design_mode)

    base_dir = common.resolve_base_directory(Path.cwd())
    if base_dir == Path.cwd() and common.path_in_appdata(base_dir):
        common.exit_with_error(
            '[ERROR] Template path was not provided. Run the uninstaller from "1. Pin templates..." so the correct folder is passed in.',
            design_mode,
        )

    _print_intro(base_dir, design_mode)

    if design_mode and common.DESIGN_LOG_UNINSTALLER:
        logging.getLogger(__name__).info("[INFO] Uninstalling from: %s", base_dir)

    destinations = common.default_destinations()
    if design_mode and common.DESIGN_LOG_UNINSTALLER:
        logging.getLogger(__name__).info(
            "[INFO] Default paths: WORD=%s PPT=%s EXCEL=%s",
            destinations.get("WORD"),
            destinations.get("POWERPOINT"),
            destinations.get("EXCEL"),
        )
    common.log_template_folder_contents(common.resolve_template_paths(), design_mode)
    common.remove_normal_templates(design_mode)
    common.remove_installed_templates(destinations, design_mode, base_dir)
    common.delete_custom_copies(base_dir, destinations, design_mode)
    common.clear_mru_entries_for_payload(base_dir, destinations, design_mode)
    common.remove_normal_templates(design_mode)
    _run_post_uninstall_actions(base_dir, design_mode)

    if design_mode and common.DESIGN_LOG_UNINSTALLER:
        logging.getLogger(__name__).info("[FINAL] Uninstall completed.")
    elif not design_mode:
        print("Ready")
    return 0


def _print_intro(base_dir: Path, design_mode: bool) -> None:
    if design_mode and common.DESIGN_LOG_UNINSTALLER:
        logging.getLogger(__name__).info("[DEBUG] Design mode enabled=true")
        logging.getLogger(__name__).info("[INFO] Base folder: %s", base_dir)
    else:
        print("Removing custom templates and restoring the Microsoft Office default settings...")


def _resolve_design_mode() -> bool:
    if MANUAL_IS_DESIGN_MODE is not None:
        return bool(MANUAL_IS_DESIGN_MODE)
    return bool(common.DEFAULT_DESIGN_MODE)


def _run_post_uninstall_actions(base_dir: Path, design_mode: bool) -> None:
    import office_files_copy_allowed_apps
    import office_files_copy_allowed_destinations

    try:
        office_files_copy_allowed_destinations.run_actions(base_dir, design_mode)
        office_files_copy_allowed_apps.run_actions(base_dir, design_mode)
    except OSError as exc:
        if design_mode and common.DESIGN_LOG_UNINSTALLER:
            logging.getLogger(__name__).warning(
                "[WARN] Post-uninstall actions could not be executed (%s)",
                exc,
            )
        elif not design_mode:
            print(f"[WARN] Post-uninstall actions could not be executed ({exc})")


if __name__ == "__main__":
    raise SystemExit(main())
