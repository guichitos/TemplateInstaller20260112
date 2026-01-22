"""Shared helpers for installing/uninstalling Office templates."""
from __future__ import annotations

import logging
import os
import shutil
import subprocess
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable, Optional, Set

from author_validation import (
    AUTHOR_VALIDATION_ENABLED,
    DEFAULT_ALLOWED_TEMPLATE_AUTHORS,
    SUPPORTED_TEMPLATE_EXTENSIONS,
    check_template_author,
    iter_template_files,
)


def normalize_path(path: Path | str | None) -> Path:
    if path is None:
        return Path()
    return Path(str(path).strip().rstrip("\\/"))


sys.path.append(str(Path(__file__).resolve().parent))
import path_utils  # type: ignore  # noqa: E402

try:
    import winreg  # type: ignore[import-not-found]
except Exception:  # pragma: no cover - non-Windows environments
    winreg = None  # type: ignore[assignment]

LOGGER = logging.getLogger(__name__)

# --------------------------------------------------------------------------- #
# Manual design flags (override environment variables)
# --------------------------------------------------------------------------- #
# Set to True/False to force logs by category; leave None to use the
# corresponding environment variable or, if missing, IsDesignModeEnabled.
MANUAL_DESIGN_LOG_PATHS: bool | None = False
MANUAL_DESIGN_LOG_MRU: bool | None = False
MANUAL_DESIGN_LOG_AUTHOR: bool | None = False
MANUAL_DESIGN_LOG_COPY_BASE: bool | None = False
MANUAL_DESIGN_LOG_COPY_CUSTOM: bool | None = False
MANUAL_DESIGN_LOG_BACKUP: bool | None = False
MANUAL_DESIGN_LOG_CLOSE_APPS: bool | None = False
MANUAL_DESIGN_LOG_INSTALLER: bool | None = False
MANUAL_DESIGN_LOG_UNINSTALLER: bool | None = False


# --------------------------------------------------------------------------- #
# Constantes base
# --------------------------------------------------------------------------- #

_BASE_PATHS = path_utils.resolve_base_paths()
APPDATA_PATH = _BASE_PATHS["APPDATA"]
DOCUMENTS_PATH = _BASE_PATHS["DOCUMENTS"]

DEFAULT_DESIGN_MODE = os.environ.get("IsDesignModeEnabled", "false").lower() == "true"
MRU_VALUE_PREFIX = "[F00000000][T01ED6D7E58D00000][O00000000]*"


def _design_flag(env_var: str, manual_override: bool | None, fallback: bool) -> bool:
    if manual_override is not None:
        return bool(manual_override)
    raw = os.environ.get(env_var)
    if raw is None:
        return fallback
    return raw.lower() == "true"


DESIGN_LOG_PATHS = _design_flag("DesignLogPaths", MANUAL_DESIGN_LOG_PATHS, DEFAULT_DESIGN_MODE)
DESIGN_LOG_MRU = _design_flag("DesignLogMRU", MANUAL_DESIGN_LOG_MRU, DEFAULT_DESIGN_MODE)
DESIGN_LOG_AUTHOR = _design_flag("DesignLogAuthor", MANUAL_DESIGN_LOG_AUTHOR, DEFAULT_DESIGN_MODE)
DESIGN_LOG_COPY_BASE = _design_flag("DesignLogCopyBase", MANUAL_DESIGN_LOG_COPY_BASE, DEFAULT_DESIGN_MODE)
DESIGN_LOG_COPY_CUSTOM = _design_flag("DesignLogCopyCustom", MANUAL_DESIGN_LOG_COPY_CUSTOM, DEFAULT_DESIGN_MODE)
DESIGN_LOG_BACKUP = _design_flag("DesignLogBackup", MANUAL_DESIGN_LOG_BACKUP, DEFAULT_DESIGN_MODE)
DESIGN_LOG_CLOSE_APPS = _design_flag("DesignLogCloseApps", MANUAL_DESIGN_LOG_CLOSE_APPS, DEFAULT_DESIGN_MODE)
DESIGN_LOG_INSTALLER = _design_flag("DesignLogInstaller", MANUAL_DESIGN_LOG_INSTALLER, DEFAULT_DESIGN_MODE)
DESIGN_LOG_UNINSTALLER = _design_flag("DesignLogUninstaller", MANUAL_DESIGN_LOG_UNINSTALLER, DEFAULT_DESIGN_MODE)


def refresh_design_log_flags(effective_design_mode: bool) -> None:
    """Update design flags for this run based on the effective mode."""
    global DESIGN_LOG_PATHS, DESIGN_LOG_MRU
    global DESIGN_LOG_AUTHOR, DESIGN_LOG_COPY_BASE, DESIGN_LOG_COPY_CUSTOM, DESIGN_LOG_BACKUP
    global DESIGN_LOG_CLOSE_APPS, DESIGN_LOG_INSTALLER, DESIGN_LOG_UNINSTALLER

    DESIGN_LOG_PATHS = _design_flag("DesignLogPaths", MANUAL_DESIGN_LOG_PATHS, effective_design_mode)
    DESIGN_LOG_MRU = _design_flag("DesignLogMRU", MANUAL_DESIGN_LOG_MRU, effective_design_mode)
    DESIGN_LOG_AUTHOR = _design_flag("DesignLogAuthor", MANUAL_DESIGN_LOG_AUTHOR, effective_design_mode)
    DESIGN_LOG_COPY_BASE = _design_flag("DesignLogCopyBase", MANUAL_DESIGN_LOG_COPY_BASE, effective_design_mode)
    DESIGN_LOG_COPY_CUSTOM = _design_flag("DesignLogCopyCustom", MANUAL_DESIGN_LOG_COPY_CUSTOM, effective_design_mode)
    DESIGN_LOG_BACKUP = _design_flag("DesignLogBackup", MANUAL_DESIGN_LOG_BACKUP, effective_design_mode)
    DESIGN_LOG_CLOSE_APPS = _design_flag("DesignLogCloseApps", MANUAL_DESIGN_LOG_CLOSE_APPS, effective_design_mode)
    DESIGN_LOG_INSTALLER = _design_flag("DesignLogInstaller", MANUAL_DESIGN_LOG_INSTALLER, effective_design_mode)
    DESIGN_LOG_UNINSTALLER = _design_flag("DesignLogUninstaller", MANUAL_DESIGN_LOG_UNINSTALLER, effective_design_mode)

DEFAULT_CUSTOM_OFFICE_TEMPLATE_PATH = normalize_path(
    os.environ.get("CUSTOM_OFFICE_TEMPLATE_PATH", _BASE_PATHS["CUSTOM_WORD"])
)
DEFAULT_POWERPOINT_TEMPLATE_PATH = normalize_path(
    os.environ.get("POWERPOINT_TEMPLATE_PATH", _BASE_PATHS["CUSTOM_PPT"])
)
DEFAULT_EXCEL_TEMPLATE_PATH = normalize_path(
    os.environ.get("EXCEL_TEMPLATE_PATH", _BASE_PATHS["CUSTOM_EXCEL"])
)
DEFAULT_ROAMING_TEMPLATE_FOLDER = normalize_path(
    os.environ.get("ROAMING_TEMPLATE_FOLDER_PATH", _BASE_PATHS["ROAMING"])
)
DEFAULT_EXCEL_STARTUP_FOLDER = normalize_path(
    os.environ.get("EXCEL_STARTUP_FOLDER_PATH", _BASE_PATHS["EXCEL_STARTUP"])
)
DEFAULT_THEME_FOLDER = normalize_path(_BASE_PATHS["THEME"])

BASE_TEMPLATE_NAMES = {
    "Normal.dotx",
    "Normal.dotm",
    "NormalEmail.dotx",
    "NormalEmail.dotm",
    "Blank.potx",
    "Blank.potm",
    "Book.xltx",
    "Book.xltm",
    "Sheet.xltx",
    "Sheet.xltm",
}


# --------------------------------------------------------------------------- #
# Generic helpers
# --------------------------------------------------------------------------- #


def ensure_directory(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def resolve_base_directory(base_dir: Path) -> Path:
    """Use only the current path as the base for templates."""
    return normalize_path(base_dir)


def path_in_appdata(path: Path) -> bool:
    try:
        return normalize_path(path).resolve().as_posix().startswith(
            normalize_path(APPDATA_PATH).resolve().as_posix()
        )
    except OSError:
        return False


def ensure_parents_and_copy(source: Path, destination: Path) -> None:
    ensure_directory(destination.parent)
    shutil.copy2(source, destination)


def _design_log(enabled: bool, design_mode: bool, level: int, message: str, *args: object) -> None:
    if design_mode and enabled:
        LOGGER.log(level, message, *args)




# --------------------------------------------------------------------------- #
# Installation / uninstallation
# --------------------------------------------------------------------------- #


@dataclass
class InstallFlags:
    totals: dict[str, int] = field(default_factory=lambda: {"files": 0, "errors": 0, "blocked": 0})


def install_template(
    app_label: str,
    filename: str,
    source_root: Path,
    destination_root: Path,
    destinations_map: dict[str, Path],
    flags: InstallFlags,
    allowed_authors: Iterable[str],
    validation_enabled: bool,
    design_mode: bool,
) -> None:
    source = normalize_path(source_root / filename)
    destination_root = ensure_directory(normalize_path(destination_root))
    destination = destination_root / filename

    if not source.exists():
        _design_log(DESIGN_LOG_COPY_BASE, design_mode, logging.WARNING, "[WARNING] Source file not found: %s", source)
        flags.totals["errors"] += 1
        return

    author_check = check_template_author(
        source,
        allowed_authors=allowed_authors,
        validation_enabled=validation_enabled,
        design_mode=design_mode,
        log_callback=lambda level, message, *args: _design_log(
            DESIGN_LOG_AUTHOR,
            design_mode,
            level,
            message,
            *args,
        ),
    )
    if not author_check.allowed:
        _design_log(DESIGN_LOG_AUTHOR, design_mode, logging.WARNING, author_check.message)
        flags.totals["blocked"] += 1
        return

    backup_existing(destination, design_mode)
    try:
        ensure_parents_and_copy(source, destination)
        flags.totals["files"] += 1
        _design_log(DESIGN_LOG_COPY_BASE, design_mode, logging.INFO, "[OK] Copied %s to %s", filename, destination)
        _update_mru_if_applicable(app_label, destination, design_mode)
    except OSError as exc:
        flags.totals["errors"] += 1
        _design_log(DESIGN_LOG_COPY_BASE, design_mode, logging.ERROR, "[ERROR] Copy failed for %s (%s)", filename, exc)
        return


def copy_custom_templates(base_dir: Path, destinations: dict[str, Path], flags: InstallFlags, allowed: Iterable[str], validation_enabled: bool, design_mode: bool) -> None:
    for file in iter_template_files(base_dir):
        filename = file.name
        extension = file.suffix.lower()
        if filename in BASE_TEMPLATE_NAMES:
            continue
        if extension in {".xltx", ".xltm"}:
            destination_root = destinations["EXCEL_CUSTOM"]
        elif extension in {".dotx", ".dotm"}:
            destination_root = destinations["WORD_CUSTOM"]
        elif extension in {".potx", ".potm"}:
            destination_root = destinations["POWERPOINT_CUSTOM"]
        else:
            destination_root = _destination_for_extension(extension, destinations)
        if destination_root is None:
            _design_log(DESIGN_LOG_COPY_CUSTOM, design_mode, logging.WARNING, "[WARNING] No destination for %s", filename)
            continue

        result = check_template_author(
            file,
            allowed_authors=allowed,
            validation_enabled=validation_enabled,
            design_mode=design_mode,
            log_callback=lambda level, message, *args: _design_log(
                DESIGN_LOG_AUTHOR,
                design_mode,
                level,
                message,
                *args,
            ),
        )
        if not result.allowed:
            flags.totals["blocked"] += 1
            _design_log(DESIGN_LOG_AUTHOR, design_mode, logging.WARNING, result.message)
            continue

        target_path = destination_root / filename
        backup_existing(target_path, design_mode)
        try:
            ensure_parents_and_copy(file, target_path)
            flags.totals["files"] += 1
            _design_log(
                DESIGN_LOG_COPY_CUSTOM,
                design_mode,
                logging.INFO,
                "[OK] Copied %s to %s",
                filename,
                target_path,
            )
            _update_mru_if_applicable_extension(extension, target_path, design_mode)
        except OSError as exc:
            flags.totals["errors"] += 1
            _design_log(DESIGN_LOG_COPY_CUSTOM, design_mode, logging.ERROR, "[ERROR] Copy failed for %s (%s)", filename, exc)
            continue


def remove_installed_templates(
    destinations: dict[str, Path],
    design_mode: bool,
    payload_dir: Path | None = None,
) -> None:
    targets = {
        destinations["WORD"]: ["Normal.dotx", "Normal.dotm", "NormalEmail.dotx", "NormalEmail.dotm"],
        destinations["POWERPOINT"]: ["Blank.potx", "Blank.potm"],
        destinations["EXCEL"]: ["Book.xltx", "Book.xltm", "Sheet.xltx", "Sheet.xltm"],
        destinations["THEMES"]: [],
    }
    failures: list[Path] = []
    for root, files in targets.items():
        for name in files:
            target = normalize_path(root / name)
            try:
                _design_log(
                    DESIGN_LOG_UNINSTALLER,
                    design_mode,
                    logging.INFO,
                    "[INFO] Checking %s",
                    target,
                )
                if not target.exists():
                    _design_log(DESIGN_LOG_UNINSTALLER, design_mode, logging.INFO, "[INFO] Does not exist %s", target)
                    continue
                backup_existing(target, design_mode)
                _design_log(DESIGN_LOG_UNINSTALLER, design_mode, logging.INFO, "[INFO] Deleting %s", target)
                target.unlink()
                _design_log(DESIGN_LOG_UNINSTALLER, design_mode, logging.INFO, "[INFO] Deleted %s", target)
                if target.exists():
                    _design_log(
                        DESIGN_LOG_UNINSTALLER,
                        design_mode,
                        logging.WARNING,
                        "[WARN] File persisted after deletion: %s",
                        target,
                    )
                    failures.append(target)
            except OSError as exc:
                _design_log(DESIGN_LOG_UNINSTALLER, design_mode, logging.WARNING, "[WARN] Could not delete %s (%s)", target, exc)
                failures.append(target)
    if failures:
        summary = ", ".join(str(path) for path in failures)
        _design_log(
            DESIGN_LOG_UNINSTALLER,
            design_mode,
            logging.WARNING,
            "[WARN] Files remained after deletion. Close Office/Outlook and try again: %s",
            summary,
        )


def remove_normal_templates(
    design_mode: bool,
    emit: Callable[[str], None] | None = None,
) -> None:
    if emit is None:
        emit = lambda message: _design_log(DESIGN_LOG_UNINSTALLER, design_mode, logging.INFO, message)
    template_dir = resolve_template_paths()["ROAMING"]
    emit('[INFO] Path retrieved from common.resolve_template_paths()["ROAMING"]')
    emit(f"[INFO] Template path (ROAMING): {template_dir}")
    if not template_dir.exists():
        emit(f"[ERROR] Folder does not exist: {template_dir}")
        return
    targets = ("Normal.dotx", "Normal.dotm", "NormalEmail.dotx", "NormalEmail.dotm")
    for filename in targets:
        target = template_dir / filename
        if not target.exists():
            emit(f"[SKIP] Does not exist: {target}")
            continue
        try:
            target.unlink()
            if target.exists():
                emit(f"[WARN] Persisted after deletion: {target}")
            else:
                emit(f"[OK] Deleted: {target}")
        except OSError as exc:
            emit(f"[ERROR] Could not delete {target} ({exc})")


def delete_custom_copies(
    base_dir: Path,
    destinations: dict[str, Path],
    design_mode: bool,
) -> None:
    for file in iter_template_files(base_dir):
        if file.name in BASE_TEMPLATE_NAMES:
            continue
        extension = file.suffix.lower()
        for dest in destinations.values():
            candidate = normalize_path(dest / file.name)
            try:
                if candidate.exists():
                    if design_mode:
                        print(f"[DELETE] Deleting file: {candidate}")
                    candidate.unlink()
                    _design_log(DESIGN_LOG_UNINSTALLER, design_mode, logging.INFO, "[INFO] Deleted %s", candidate)
            except OSError as exc:
                _design_log(DESIGN_LOG_UNINSTALLER, design_mode, logging.WARNING, "[WARN] Could not delete %s (%s)", candidate, exc)


def clear_mru_entries_for_payload(base_dir: Path, destinations: dict[str, Path], design_mode: bool) -> None:
    """Remove MRU entries for payload templates and base templates."""
    if not is_windows() or winreg is None:
        return
    targets = _collect_mru_targets(base_dir, destinations)
    if not targets:
        return
    grouped: dict[str, set[str]] = {"WORD": set(), "POWERPOINT": set(), "EXCEL": set()}
    for path in targets:
        ext = path.suffix.lower()
        if ext in {".dotx", ".dotm"}:
            grouped["WORD"].add(str(path))
        elif ext in {".potx", ".potm"}:
            grouped["POWERPOINT"].add(str(path))
        elif ext in {".xltx", ".xltm"}:
            grouped["EXCEL"].add(str(path))
    for app_label, paths in grouped.items():
        if paths:
            _clear_mru_for_app(app_label, paths, design_mode)


def backup_existing(target_file: Path, design_mode: bool) -> None:
    if not target_file.exists():
        return
    backup_dir = target_file.parent / "Backups"
    ensure_directory(backup_dir)
    timestamp = datetime.now().strftime("%Y.%m.%d.%H%M")
    backup_path = backup_dir / f"{timestamp} - {target_file.name}"
    try:
        shutil.copy2(target_file, backup_path)
        _design_log(DESIGN_LOG_BACKUP, design_mode, logging.INFO, "[BACKUP] Copy created at %s", backup_path)
    except OSError as exc:
        _design_log(
            DESIGN_LOG_BACKUP,
            design_mode,
            logging.WARNING,
            "[WARN] Could not create backup of %s (%s)",
            target_file,
            exc,
        )




def _update_mru_if_applicable(app_label: str, destination: Path, design_mode: bool) -> None:
    if not _should_update_mru(destination):
        return
    ext = destination.suffix.lower()
    if ext in {".dotx", ".dotm", ".potx", ".potm", ".xltx", ".xltm"}:
        update_mru_for_template(app_label, destination, design_mode)


def _update_mru_if_applicable_extension(extension: str, destination: Path, design_mode: bool) -> None:
    if not _should_update_mru(destination):
        return
    if extension in {".dotx", ".dotm"}:
        update_mru_for_template("WORD", destination, design_mode)
    if extension in {".potx", ".potm"}:
        update_mru_for_template("POWERPOINT", destination, design_mode)
    if extension in {".xltx", ".xltm"}:
        update_mru_for_template("EXCEL", destination, design_mode)


def _should_update_mru(path: Path) -> bool:
    name = path.name
    ext = path.suffix.lower()
    if name in BASE_TEMPLATE_NAMES:
        return False
    if ext == ".thmx":
        return False
    return True


def _collect_mru_targets(base_dir: Path, destinations: dict[str, Path]) -> list[Path]:
    """Return potential MRU paths to clear (base + custom payload)."""
    targets: set[Path] = set()
    # Base templates
    base_targets = {
        "WORD": ["Normal.dotx", "Normal.dotm", "NormalEmail.dotx", "NormalEmail.dotm"],
        "POWERPOINT": ["Blank.potx", "Blank.potm"],
        "EXCEL": ["Book.xltx", "Book.xltm", "Sheet.xltx", "Sheet.xltm"],
    }
    for app, names in base_targets.items():
        dest = destinations.get(app)
        if dest:
            for name in names:
                targets.add(normalize_path(dest / name))
    # Custom payload templates
    for file in iter_template_files(base_dir):
        if file.name in BASE_TEMPLATE_NAMES:
            continue
        ext = file.suffix.lower()
        if ext == ".thmx":
            continue
        if ext in {".dotx", ".dotm"}:
            dest = destinations.get("WORD_CUSTOM")
        elif ext in {".potx", ".potm"}:
            dest = destinations.get("POWERPOINT_CUSTOM")
        elif ext in {".xltx", ".xltm"}:
            dest = destinations.get("EXCEL_CUSTOM")
        else:
            dest = _destination_for_extension(ext, destinations)
        if dest:
            targets.add(normalize_path(dest / file.name))
    return list(targets)


def _clear_mru_for_app(app_label: str, target_paths: Set[str], design_mode: bool) -> None:
    mru_paths = _find_mru_paths(app_label)
    if design_mode and DESIGN_LOG_MRU:
        LOGGER.info("[MRU] Cleanup for %s, target paths=%s", app_label, sorted(target_paths))
    for mru_path in mru_paths:
        try:
            _rewrite_mru_excluding(mru_path, target_paths, design_mode)
        except OSError as exc:
            _design_log(DESIGN_LOG_MRU, design_mode, logging.WARNING, "[MRU] Could not clean %s (%s)", mru_path, exc)


# --------------------------------------------------------------------------- #
# Platform utilities
# --------------------------------------------------------------------------- #


def is_windows() -> bool:
    return os.name == "nt"


def close_office_apps(design_mode: bool) -> None:
    if not is_windows():
        return
    processes = ("WINWORD.EXE", "POWERPNT.EXE", "EXCEL.EXE", "OUTLOOK.EXE")
    for exe in processes:
        try:
            os.system(f"taskkill /IM {exe} /F >nul 2>&1")
        except OSError:
            _design_log(DESIGN_LOG_CLOSE_APPS, design_mode, logging.DEBUG, "[DEBUG] Could not close %s", exe)
    for exe in processes:
        try:
            result = subprocess.run(
                ["tasklist", "/FI", f"IMAGENAME eq {exe}", "/NH"],
                capture_output=True,
                text=True,
            )
            output = (result.stdout or "") + (result.stderr or "")
            if exe.lower() in output.lower():
                os.system(f"taskkill /IM {exe} /F >nul 2>&1")
        except OSError:
            _design_log(DESIGN_LOG_CLOSE_APPS, design_mode, logging.DEBUG, "[DEBUG] Could not verify %s", exe)




# --------------------------------------------------------------------------- #
# Path utilities
# --------------------------------------------------------------------------- #


def default_destinations() -> dict[str, Path]:
    paths = resolve_template_paths()
    return {
        "WORD": paths["ROAMING"],
        "POWERPOINT": paths["ROAMING"],
        "EXCEL": paths["EXCEL"],
        "CUSTOM": paths["CUSTOM_WORD"],
        "WORD_CUSTOM": paths["CUSTOM_WORD"],
        "POWERPOINT_CUSTOM": paths["CUSTOM_PPT"],
        "EXCEL_CUSTOM": paths["CUSTOM_EXCEL"],
        "ROAMING": paths["ROAMING"],
        "THEMES": paths["THEME"],
    }


def resolve_template_paths() -> dict[str, Path]:
    return {
        "THEME": DEFAULT_THEME_FOLDER,
        "CUSTOM_WORD": DEFAULT_CUSTOM_OFFICE_TEMPLATE_PATH,
        "CUSTOM_PPT": DEFAULT_POWERPOINT_TEMPLATE_PATH or DEFAULT_CUSTOM_OFFICE_TEMPLATE_PATH,
        "CUSTOM_EXCEL": DEFAULT_EXCEL_TEMPLATE_PATH or DEFAULT_CUSTOM_OFFICE_TEMPLATE_PATH,
        "ROAMING": DEFAULT_ROAMING_TEMPLATE_FOLDER,
        "EXCEL": DEFAULT_EXCEL_STARTUP_FOLDER,
    }


def log_template_paths(paths: dict[str, Path], design_mode: bool) -> None:
    if not design_mode or not DESIGN_LOG_PATHS:
        return
    logger = logging.getLogger(__name__)
    logger.info("================= CALCULATED PATHS =================")
    logger.info("THEME_PATH                  = %s", paths["THEME"])
    logger.info("CUSTOM_WORD_TEMPLATE_PATH   = %s", paths["CUSTOM_WORD"])
    logger.info("CUSTOM_PPT_TEMPLATE_PATH    = %s", paths["CUSTOM_PPT"])
    logger.info("CUSTOM_EXCEL_TEMPLATE_PATH  = %s", paths["CUSTOM_EXCEL"])
    logger.info("ROAMING_TEMPLATE_PATH       = %s", paths["ROAMING"])
    logger.info("EXCEL_STARTUP_PATH          = %s", paths["EXCEL"])
    logger.info("====================================================")


def log_template_folder_contents(paths: dict[str, Path], design_mode: bool) -> None:
    if not design_mode or not (DESIGN_LOG_PATHS or DESIGN_LOG_MRU):
        return
    logger = logging.getLogger(__name__)
    targets = [
        ("THEME_PATH", paths["THEME"]),
        ("CUSTOM_WORD_TEMPLATE_PATH", paths["CUSTOM_WORD"]),
        ("CUSTOM_PPT_TEMPLATE_PATH", paths["CUSTOM_PPT"]),
        ("CUSTOM_EXCEL_TEMPLATE_PATH", paths["CUSTOM_EXCEL"]),
        ("ROAMING_TEMPLATE_PATH", paths["ROAMING"]),
        ("EXCEL_STARTUP_PATH", paths["EXCEL"]),
    ]
    for label, folder in targets:
        try:
            if not folder.exists():
                logger.info("[INFO] %s does not exist: %s", label, folder)
                continue
            files = sorted(p.name for p in folder.iterdir() if p.is_file())
            logger.info("[INFO] %s (%s): %s", label, folder, ", ".join(files) if files else "[empty]")
        except OSError as exc:
            logger.warning("[WARN] Could not list %s (%s)", folder, exc)


def log_registry_sources(design_mode: bool) -> None:
    if not design_mode or not DESIGN_LOG_MRU:
        return
    logger = logging.getLogger(__name__)
    word_personal = path_utils.read_registry_value(
        r"Software\Microsoft\Office\\16.0\\Word\\Options",
        "PersonalTemplates",
    )
    word_user = path_utils.read_registry_value(
        r"Software\Microsoft\Office\\16.0\\Common\\General",
        "UserTemplates",
    )
    ppt_personal = path_utils.read_registry_value(
        r"Software\Microsoft\Office\\16.0\\PowerPoint\\Options",
        "PersonalTemplates",
    )
    ppt_user = path_utils.read_registry_value(
        r"Software\Microsoft\Office\\16.0\\Common\\General",
        "UserTemplates",
    )
    excel_personal = path_utils.read_registry_value(
        r"Software\Microsoft\Office\\16.0\\Excel\\Options",
        "PersonalTemplates",
    )
    excel_user = path_utils.read_registry_value(
        r"Software\Microsoft\Office\\16.0\\Common\\General",
        "UserTemplates",
    )
    logger.info("[REG] Word PersonalTemplates: %s", word_personal or "[no value]")
    logger.info("[REG] Word UserTemplates: %s", word_user or "[no value]")
    logger.info("[REG] PowerPoint PersonalTemplates: %s", ppt_personal or "[no value]")
    logger.info("[REG] PowerPoint UserTemplates: %s", ppt_user or "[no value]")
    logger.info("[REG] Excel PersonalTemplates: %s", excel_personal or "[no value]")
    logger.info("[REG] Excel UserTemplates: %s", excel_user or "[no value]")


def update_mru_for_template(app_label: str, file_path: Path, design_mode: bool) -> None:
    if not is_windows() or winreg is None:
        return
    mru_paths = _find_mru_paths(app_label)
    if design_mode and DESIGN_LOG_MRU:
        LOGGER.info("[MRU] Updating MRU for %s at paths: %s", app_label, mru_paths)
    for mru_path in mru_paths:
        try:
            _write_mru_entry(mru_path, file_path, design_mode)
        except OSError as exc:
            if design_mode and DESIGN_LOG_MRU:
                LOGGER.warning("[MRU] Could not write to %s (%s)", mru_path, exc)


def _find_mru_paths(app_label: str) -> list[str]:
    reg_name = _app_registry_name(app_label)
    if not reg_name:
        return []
    roots: list[str] = []
    versions = ("16.0", "15.0", "14.0", "12.0")
    for version in versions:
        base = fr"Software\Microsoft\Office\{version}\{reg_name}\Recent Templates"
        # Prefer LiveID/ADAL containers if they exist
        if winreg:
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, base) as root:
                    sub_count = winreg.QueryInfoKey(root)[0]
                    for idx in range(sub_count):
                        sub = winreg.EnumKey(root, idx)
                        if sub.upper().startswith("ADAL_") or sub.upper().startswith("LIVEID_"):
                            roots.append(f"HKCU\\{base}\\{sub}\\File MRU")
            except OSError:
                pass
        roots.append(f"HKCU\\{base}\\File MRU")
    # Deduplicate while preserving order
    seen: set[str] = set()
    ordered: list[str] = []
    for path in roots:
        if path not in seen:
            seen.add(path)
            ordered.append(path)
    return ordered


def _app_registry_name(app_label: str) -> str:
    mapping = {"WORD": "Word", "POWERPOINT": "PowerPoint", "EXCEL": "Excel"}
    return mapping.get(app_label.upper(), "")


def _write_mru_entry(reg_path: str, file_path: Path, design_mode: bool) -> None:
    if winreg is None:
        return
    file_path = normalize_path(file_path)
    full_path = str(file_path)
    basename = file_path.stem
    hive, subkey = reg_path.split("\\", 1)
    hive_obj = winreg.HKEY_CURRENT_USER if hive.upper() == "HKCU" else None
    if hive_obj is None:
        return
    try:
        key = winreg.CreateKeyEx(hive_obj, subkey, 0, winreg.KEY_ALL_ACCESS)
    except OSError:
        return
    with key:
        # Read existing entries
        existing_items: list[tuple[int, str]] = []
        index = 0
        try:
            while True:
                name, value, _ = winreg.EnumValue(key, index)
                if name.startswith("Item Metadata "):
                    index += 1
                    continue
                if name.startswith("Item "):
                    try:
                        num = int(name.split(" ", 1)[1])
                    except Exception:
                        num = 0
                    if isinstance(value, str):
                        extracted = _extract_mru_path(value)
                        if extracted:
                            existing_items.append((num, extracted))
                index += 1
        except OSError:
            pass
        # Filter duplicates for the same path
        filtered = []
        for _, value in existing_items:
            if full_path.lower() == value.lower():
                continue
            filtered.append(value)
        # Prepare new list with the file at the front
        new_entries: list[str] = [full_path] + filtered
        # Limit to e.g. 10 entries
        new_entries = new_entries[:10]
        # Rewrite
        for idx, entry in enumerate(new_entries, start=1):
            item_name = f"Item {idx}"
            meta_name = f"Item Metadata {idx}"
            reg_value = f"{MRU_VALUE_PREFIX}{entry}"
            meta_value = f"<Metadata><AppSpecific><id>{entry}</id><nm>{basename}</nm><du>{entry}</du></AppSpecific></Metadata>"
            _design_log(DESIGN_LOG_MRU, design_mode, logging.INFO, "[MRU] %s -> %s", item_name, entry)
            _design_log(DESIGN_LOG_MRU, design_mode, logging.DEBUG, "[MRU] %s (name=%s)", meta_name, basename)
            winreg.SetValueEx(key, item_name, 0, winreg.REG_SZ, reg_value)
            winreg.SetValueEx(key, meta_name, 0, winreg.REG_SZ, meta_value)
        _design_log(DESIGN_LOG_MRU, design_mode, logging.INFO, "[MRU] %s updated with %s", reg_path, full_path)


def _extract_mru_path(raw_value: str) -> Optional[str]:
    if not raw_value:
        return None
    if "*" in raw_value:
        candidate = raw_value.split("*")[-1]
        return candidate.strip() or None
    return raw_value.strip() or None


def _rewrite_mru_excluding(mru_path: str, targets: Set[str], design_mode: bool) -> None:
    """Rewrite the MRU excluding target paths and reindex items."""
    if winreg is None:
        return
    hive, subkey = mru_path.split("\\", 1)
    hive_obj = winreg.HKEY_CURRENT_USER if hive.upper() == "HKCU" else None
    if hive_obj is None:
        return
    try:
        key = winreg.CreateKeyEx(hive_obj, subkey, 0, winreg.KEY_ALL_ACCESS)
    except OSError:
        return
    with key:
        items: list[tuple[int, str]] = []
        metadata: dict[int, str] = {}
        index = 0
        try:
            while True:
                name, value, _ = winreg.EnumValue(key, index)
                if not isinstance(value, str):
                    index += 1
                    continue
                if name.startswith("Item Metadata "):
                    try:
                        num = int(name.split(" ", 2)[2])
                        metadata[num] = value
                    except Exception:
                        pass
                elif name.startswith("Item "):
                    try:
                        num = int(name.split(" ", 1)[1])
                    except Exception:
                        num = 0
                    items.append((num, value))
                index += 1
        except OSError:
            pass
        # Delete everything before rewriting
        try:
            index = 0
            while True:
                name, _, _ = winreg.EnumValue(key, index)
                if name.startswith("Item"):
                    winreg.DeleteValue(key, name)
                    continue
                index += 1
        except OSError:
            pass
        # Filter and reindex
        target_lowers = {t.lower() for t in targets}
        filtered: list[tuple[str, str]] = []
        for idx_num, value in sorted(items, key=lambda x: x[0]):
            path = _extract_mru_path(value)
            if path and path.lower() in target_lowers:
                continue
            meta_val = metadata.get(idx_num, "")
            filtered.append((value, meta_val))
        for new_idx, (val, meta_val) in enumerate(filtered, start=1):
            item_name = f"Item {new_idx}"
            meta_name = f"Item Metadata {new_idx}"
            _design_log(DESIGN_LOG_MRU, design_mode, logging.INFO, "[MRU] Cleanup %s -> %s", item_name, _extract_mru_path(val) or val)
            winreg.SetValueEx(key, item_name, 0, winreg.REG_SZ, val)
            if meta_val:
                winreg.SetValueEx(key, meta_name, 0, winreg.REG_SZ, meta_val)
def _destination_for_extension(extension: str, destinations: dict[str, Path]) -> Optional[Path]:
    if extension in {".dotx", ".dotm"}:
        return destinations["WORD"]
    if extension in {".potx", ".potm"}:
        return destinations["POWERPOINT"]
    if extension in {".xltx", ".xltm"}:
        return destinations["EXCEL"]
    if extension == ".thmx":
        return destinations["THEMES"]
    return None


def configure_logging(design_mode: bool) -> None:
    level = logging.DEBUG if design_mode else logging.INFO
    logging.basicConfig(level=level, format="%(message)s")


def exit_with_error(message: str, design_mode: bool = DEFAULT_DESIGN_MODE) -> None:
    if design_mode:
        print(message)
    sys.exit(1)
