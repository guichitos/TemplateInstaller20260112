"""Single installer based on the current folder."""
from __future__ import annotations

import argparse
import logging
import os
from pathlib import Path
from typing import Iterable

# Manual configuration for design mode.
# - Set to True to force design mode on.
# - Set to False to force design mode off.
# - Leave as None to use the normal environment-based logic.
MANUAL_IS_DESIGN_MODE: bool | None = False


# Ensure access to the shared `common` module in all execution modes
try:
    from . import common
except ImportError:
    import sys

    sys.path.append(str(Path(__file__).resolve().parent))
    import common 


# Read installer options from the command line
def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Office template installer (Python)")
    parser.add_argument(
        "--allowed-authors",
        help="Semicolon-separated list of allowed authors.",
    )
    parser.add_argument(
        "--check-author",
        metavar="RUTA",
        help="Only validate the author for a file/folder and exit.",
    )
    return parser.parse_args()


# Main installer workflow
def main(argv: Iterable[str] | None = None) -> int:
    # Parse command line arguments
    args = parse_args()
    design_mode = _resolve_design_mode()
    common.refresh_design_log_flags(design_mode)    # Update log flags based on design mode
    common.configure_logging(design_mode) 

    # Resolve template paths and determine the base folder
    resolved_paths = common.resolve_template_paths()
    common.log_registry_sources(design_mode)
    common.log_template_paths(resolved_paths, design_mode)
    if design_mode and common.DESIGN_LOG_PATHS:
        logging.getLogger(__name__).info(
            "[INFO] Extra template folder (WORD): %s", resolved_paths["CUSTOM_WORD"]
        )
        logging.getLogger(__name__).info(
            "[INFO] Extra template folder (POWERPOINT): %s", resolved_paths["CUSTOM_PPT"]
        )
        logging.getLogger(__name__).info(
            "[INFO] Extra template folder (EXCEL): %s", resolved_paths["CUSTOM_EXCEL"]
        )

    # Validate base directory and author rules
    working_dir = Path.cwd()     
    base_dir = common.resolve_base_directory(working_dir)

    if base_dir == working_dir and common.path_in_appdata(working_dir):
        common.exit_with_error(
            '[ERROR] Template path was not provided. Run the installer from "1. Pin templates..." so the correct folder is passed in.',
            design_mode,
        )

    # Resolve allowed authors and validation settings
    allowed_authors = _resolve_allowed_authors(args.allowed_authors)
    validation_enabled = common.AUTHOR_VALIDATION_ENABLED

    if args.check_author:
        result = common.check_template_author(
            Path(args.check_author),
            allowed_authors=allowed_authors,
            validation_enabled=validation_enabled,
            design_mode=design_mode,
        )
        print(result.as_cli_output())
        if design_mode and common.DESIGN_LOG_AUTHOR:
            logging.getLogger(__name__).info(result.message)
        return 0 if result.allowed else 1

    _print_intro(base_dir, design_mode) # Print introductory message
    common.close_office_apps(design_mode) # Close Office applications if needed

    destinations = common.default_destinations()
    flags = common.InstallFlags()

    # Define base Office templates to install
    base_targets = [
        ("WORD", "Normal.dotx", destinations["WORD"]),
        ("WORD", "Normal.dotm", destinations["WORD"]),
        ("WORD", "NormalEmail.dotx", destinations["WORD"]),
        ("WORD", "NormalEmail.dotm", destinations["WORD"]),
        ("POWERPOINT", "Blank.potx", destinations["POWERPOINT"]),
        ("POWERPOINT", "Blank.potm", destinations["POWERPOINT"]),
        ("EXCEL", "Book.xltx", destinations["EXCEL"]),
        ("EXCEL", "Book.xltm", destinations["EXCEL"]),
        ("EXCEL", "Sheet.xltx", destinations["EXCEL"]),
        ("EXCEL", "Sheet.xltm", destinations["EXCEL"]),
    ]

    # Install base templates
    for app_label, filename, destination in base_targets:
        common.install_template(
            app_label,
            filename,
            base_dir,
            destination,
            destinations,
            flags,
            allowed_authors,
            validation_enabled,
            design_mode,
        )

    # Custom templates
    common.copy_custom_templates(
        base_dir=base_dir,
        destinations=destinations,
        flags=flags,
        allowed=allowed_authors,
        validation_enabled=validation_enabled,
        design_mode=design_mode,
    )

    # Run post-installation actions
    _run_post_install_actions(base_dir, design_mode) 

    # Final summary
    if design_mode and common.DESIGN_LOG_INSTALLER:
        logging.getLogger(__name__).info(
            "[FINAL] Installation completed. Files copied=%s, errors=%s, blocked=%s.",
            flags.totals["files"],
            flags.totals["errors"],
            flags.totals["blocked"],
        )
    elif not design_mode:
        print("Ready")
    return 0

# Print introductory message
def _print_intro(base_dir: Path, design_mode: bool) -> None:
    if design_mode and common.DESIGN_LOG_INSTALLER:
        logging.getLogger(__name__).info("[DEBUG] Design mode enabled=true")
        logging.getLogger(__name__).info("[INFO] Base folder: %s", base_dir)
    else:
        print("Installing custom templates and applying them as the new Microsoft Office defaults...")

#
def _resolve_allowed_authors(cli_value: str | None) -> list[str]:
    env_value = os.environ.get("AllowedTemplateAuthors")
    raw = cli_value or env_value
    if not raw:
        return common.DEFAULT_ALLOWED_TEMPLATE_AUTHORS
    return [author.strip() for author in raw.split(";") if author.strip()]

# Determine if design mode is enabled
def _resolve_design_mode() -> bool:
    if MANUAL_IS_DESIGN_MODE is not None:
        return bool(MANUAL_IS_DESIGN_MODE)
    return bool(common.DEFAULT_DESIGN_MODE)

# Execute post-installation actions
def _run_post_install_actions(base_dir: Path, design_mode: bool) -> None:
    import office_files_copy_allowed_apps
    import office_files_copy_allowed_destinations

    try:
        office_files_copy_allowed_destinations.run_actions(base_dir, design_mode)
        office_files_copy_allowed_apps.run_actions(base_dir, design_mode)
    except OSError as exc:
        if design_mode and common.DESIGN_LOG_INSTALLER:
            logging.getLogger(__name__).warning(
                "[WARN] Post-install actions could not be executed (%s)",
                exc,
            )
        elif not design_mode:
            print(f"[WARN] Post-install actions could not be executed ({exc})")


if __name__ == "__main__":
    raise SystemExit(main())
