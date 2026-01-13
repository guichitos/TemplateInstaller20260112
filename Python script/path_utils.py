"""Resolución de rutas de plantillas fuera de common."""
from __future__ import annotations

import importlib.util
import os
from pathlib import Path
from typing import Optional


def normalize_path(path: Path | str | None) -> Path:
    if path is None:
        return Path()
    return Path(str(path).strip().rstrip("\\/"))


_WINREG_SPEC = importlib.util.find_spec("winreg")
if _WINREG_SPEC is not None:  # pragma: no cover - Windows
    import winreg  # type: ignore[import-not-found]
else:  # pragma: no cover - no Windows
    winreg = None  # type: ignore[assignment]


def read_registry_value(path: str, name: str) -> Optional[str]:
    if winreg is None:
        return None
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:  # type: ignore[union-attr]
            value, _ = winreg.QueryValueEx(key, name)  # type: ignore[union-attr]
            return os.path.expandvars(str(value))
    except OSError:
        return None


def _resolve_appdata_path() -> Path:
    appdata = read_registry_value(
        r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
        "AppData",
    )
    if not appdata:
        appdata = os.environ.get("APPDATA")
    return normalize_path(appdata or (Path.home() / "AppData" / "Roaming"))


def _resolve_documents_path() -> Path:
    documents = read_registry_value(
        r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
        "Personal",
    )
    if not documents:
        documents = os.environ.get("USERPROFILE")
        if documents:
            documents = str(Path(documents) / "Documents")
    return normalize_path(documents or (Path.home() / "Documents"))


def _resolve_custom_template_path(default_custom_dir: Path) -> Path:
    if winreg:
        for version in ("16.0", "15.0", "14.0", "12.0"):
            value = read_registry_value(
                fr"Software\Microsoft\Office\{version}\Word\Options",
                "PersonalTemplates",
            )
            if value:
                return normalize_path(value)
        for version in ("16.0", "15.0", "14.0", "12.0"):
            value = read_registry_value(
                fr"Software\Microsoft\Office\{version}\Common\General",
                "UserTemplates",
            )
            if value:
                return normalize_path(value)
    return normalize_path(default_custom_dir)


def _resolve_custom_alt_path(custom_primary: Path, default_custom_dir: Path, default_alt_dir: Path) -> Path:
    if winreg:
        for version in ("16.0", "15.0", "14.0", "12.0"):
            value = read_registry_value(
                fr"Software\Microsoft\Office\{version}\PowerPoint\Options",
                "PersonalTemplates",
            )
            if value:
                return normalize_path(value)
        for version in ("16.0", "15.0", "14.0", "12.0"):
            value = read_registry_value(
                fr"Software\Microsoft\Office\{version}\Common\General",
                "UserTemplates",
            )
            if value:
                return normalize_path(value)
    return normalize_path(custom_primary or default_custom_dir or default_alt_dir)


def _resolve_excel_template_path(custom_primary: Path, default_custom_dir: Path, default_alt_dir: Path) -> Path:
    if winreg:
        for version in ("16.0", "15.0", "14.0", "12.0"):
            value = read_registry_value(
                fr"Software\Microsoft\Office\{version}\Excel\Options",
                "PersonalTemplates",
            )
            if value:
                return normalize_path(value)
        for version in ("16.0", "15.0", "14.0", "12.0"):
            value = read_registry_value(
                fr"Software\Microsoft\Office\{version}\Common\General",
                "UserTemplates",
            )
            if value:
                return normalize_path(value)
    return normalize_path(custom_primary or default_custom_dir or default_alt_dir)


def _log_paths_if_design_mode(paths: dict[str, Path]) -> None:
    design_mode = os.environ.get("IsDesignModeEnabled", "false").lower() == "true"
    if not design_mode:
        return
    print("[PATHS] Rutas resueltas:")
    for key, value in paths.items():
        print(f"[PATHS] {key} = {value}")


def resolve_base_paths() -> dict[str, Path]:
    documents_path = _resolve_documents_path()
    default_custom_dir = documents_path / "Custom Office Templates"
    default_custom_alt_dir = documents_path / "Plantillas personalizadas de Office"
    custom_word = _resolve_custom_template_path(default_custom_dir)
    custom_ppt = _resolve_custom_alt_path(custom_word, default_custom_dir, default_custom_alt_dir)
    custom_excel = _resolve_excel_template_path(custom_word, default_custom_dir, default_custom_alt_dir)
    appdata_path = _resolve_appdata_path()
    paths = {
        "APPDATA": appdata_path,
        "DOCUMENTS": documents_path,
        "CUSTOM_WORD": custom_word,
        "CUSTOM_PPT": custom_ppt,
        "CUSTOM_EXCEL": custom_excel,
        "CUSTOM_ADDITIONAL": default_custom_alt_dir,
        "THEME": appdata_path / "Microsoft" / "Templates" / "Document Themes",
        "ROAMING": appdata_path / "Microsoft" / "Templates",
        "EXCEL_STARTUP": appdata_path / "Microsoft" / "Excel" / "XLSTART",
    }
    _log_paths_if_design_mode(paths)
    return paths


def _print_paths() -> None:
    paths = resolve_base_paths()
    print("[PATHS] Rutas resueltas:")
    for key, value in paths.items():
        print(f"[PATHS] {key} = {value}")


if __name__ == "__main__":
    _print_paths()
