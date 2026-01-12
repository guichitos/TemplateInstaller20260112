"""Instalador único basado en la carpeta actual."""
from __future__ import annotations

import argparse
import logging
import os
import time
from pathlib import Path
from typing import Iterable

# Configuración manual para el modo diseño.
# - Establece en True para forzar modo diseño siempre.
# - Establece en False para desactivarlo siempre.
# - Deja en None para usar la lógica normal basada en entorno.
MANUAL_IS_DESIGN_MODE: bool | None = False

try:
    from . import common
except ImportError:  # pragma: no cover - permite ejecución directa como script
    import sys

    sys.path.append(str(Path(__file__).resolve().parent))
    import common  # type: ignore[no-redef]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Instalador de plantillas de Office (Python)")
    parser.add_argument(
        "--allowed-authors",
        help="Lista separada por ';' de autores permitidos.",
    )
    parser.add_argument(
        "--check-author",
        metavar="RUTA",
        help="Solo valida autor de archivo/carpeta y termina.",
    )
    return parser.parse_args()


def main(argv: Iterable[str] | None = None) -> int:
    args = parse_args()
    design_mode = _resolve_design_mode()
    common.refresh_design_log_flags(design_mode)
    common.configure_logging(design_mode)

    resolved_paths = common.resolve_template_paths()
    common.log_registry_sources(design_mode)
    common.log_template_paths(resolved_paths, design_mode)
    if design_mode and common.DESIGN_LOG_PATHS:
        logging.getLogger(__name__).info(
            "[INFO] Carpeta de plantillas extra WORD: %s", resolved_paths["CUSTOM_WORD"]
        )
        logging.getLogger(__name__).info(
            "[INFO] Carpeta de plantillas extra POWERPOINT: %s", resolved_paths["CUSTOM_PPT"]
        )
        logging.getLogger(__name__).info(
            "[INFO] Carpeta de plantillas extra EXCEL: %s", resolved_paths["CUSTOM_EXCEL"]
        )

    working_dir = Path.cwd()
    base_dir = common.resolve_base_directory(working_dir)

    if base_dir == working_dir and common.path_in_appdata(working_dir):
        common.exit_with_error(
            '[ERROR] No se recibió la ruta de las plantillas. Ejecute el instalador desde "1. Pin templates..." para que se le pase la carpeta correcta.',
            design_mode,
        )

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

    _print_intro(base_dir, design_mode)
    common.close_office_apps(design_mode)

    destinations = common.default_destinations()
    common.open_template_folders(resolved_paths, design_mode)
    flags = common.InstallFlags()

    # Plantillas base
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

    # Plantillas personalizadas
    common.copy_custom_templates(
        base_dir=base_dir,
        destinations=destinations,
        flags=flags,
        allowed=allowed_authors,
        validation_enabled=validation_enabled,
        design_mode=design_mode,
    )
    common.open_template_folders(resolved_paths, design_mode, flags)

    if flags.open_document_theme and common.DEFAULT_DOCUMENT_THEME_DELAY_SECONDS > 0:
        if design_mode and common.DESIGN_LOG_APP_LAUNCH:
            logging.getLogger(__name__).info(
                "[INFO] Esperando %s segundos antes de abrir aplicaciones...",
                common.DEFAULT_DOCUMENT_THEME_DELAY_SECONDS,
            )
        time.sleep(common.DEFAULT_DOCUMENT_THEME_DELAY_SECONDS)

    common.launch_office_apps(flags, design_mode)

    if design_mode and common.DESIGN_LOG_INSTALLER:
        logging.getLogger(__name__).info(
            "[FINAL] Instalación completada. Archivos copiados=%s, errores=%s, bloqueados=%s.",
            flags.totals["files"],
            flags.totals["errors"],
            flags.totals["blocked"],
        )
    else:
        print("Ready")
    return 0


def _print_intro(base_dir: Path, design_mode: bool) -> None:
    if design_mode and common.DESIGN_LOG_INSTALLER:
        logging.getLogger(__name__).info("[DEBUG] Modo diseño habilitado=true")
        logging.getLogger(__name__).info("[INFO] Carpeta base: %s", base_dir)
    else:
        print("Installing custom templates and applying them as the new Microsoft Office defaults...")


def _resolve_allowed_authors(cli_value: str | None) -> list[str]:
    env_value = os.environ.get("AllowedTemplateAuthors")
    raw = cli_value or env_value
    if not raw:
        return common.DEFAULT_ALLOWED_TEMPLATE_AUTHORS
    return [author.strip() for author in raw.split(";") if author.strip()]


def _resolve_design_mode() -> bool:
    if MANUAL_IS_DESIGN_MODE is not None:
        return bool(MANUAL_IS_DESIGN_MODE)
    return bool(common.DEFAULT_DESIGN_MODE)


if __name__ == "__main__":
    raise SystemExit(main())
