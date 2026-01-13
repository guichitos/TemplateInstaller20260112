"""Validación de autores en plantillas de Office."""
from __future__ import annotations

import logging
import os
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable, Iterator, List, Optional
import xml.etree.ElementTree as ET


SUPPORTED_TEMPLATE_EXTENSIONS = {
    ".dotx",
    ".dotm",
    ".potx",
    ".potm",
    ".xltx",
    ".xltm",
    ".thmx",
}

DEFAULT_ALLOWED_TEMPLATE_AUTHORS = [
    "www.grada.cc",
    "www.gradaz.com",
]

AUTHOR_VALIDATION_ENABLED = os.environ.get("AuthorValidationEnabled", "TRUE").lower() != "false"


def normalize_path(path: Path | str | None) -> Path:
    if path is None:
        return Path()
    return Path(str(path).strip().rstrip("\\/"))


def iter_template_files(base_dir: Path) -> Iterator[Path]:
    for ext in SUPPORTED_TEMPLATE_EXTENSIONS:
        yield from base_dir.glob(f"*{ext}")


@dataclass
class AuthorCheckResult:
    allowed: bool
    message: str
    authors: List[str]
    error: bool = False

    def as_cli_output(self) -> str:
        return "TRUE" if self.allowed and not self.error else "FALSE"


def check_template_author(
    target: Path,
    allowed_authors: Iterable[str] | None = None,
    validation_enabled: bool = True,
    design_mode: bool = False,
    log_callback: Callable[[int, str, object], None] | None = None,
) -> AuthorCheckResult:
    allowed = _normalize_allowed_authors(allowed_authors or DEFAULT_ALLOWED_TEMPLATE_AUTHORS)
    target = normalize_path(target)

    def _log(level: int, message: str, *args: object) -> None:
        if log_callback:
            log_callback(level, message, *args)

    if not target.exists():
        return AuthorCheckResult(
            allowed=False,
            message=f"[ERROR] No se encontró la ruta: \"{target}\"",
            authors=[],
            error=True,
        )

    if target.is_dir():
        authors_found: list[str] = []
        for file in iter_template_files(target):
            if file.suffix.lower() == ".thmx":
                _log(logging.INFO, "Archivo: %s - Autor: [OMITIDO TEMA]", file.name)
                continue
            author, error = _extract_author(file)
            if error:
                _log(logging.WARNING, error)
            if author:
                authors_found.append(author)
                _log(logging.INFO, "Archivo: %s - Autor: %s", file.name, author)
            else:
                _log(logging.INFO, "Archivo: %s - Autor: [VACÍO]", file.name)

        message = (
            f"[INFO] Autores listados para la carpeta \"{target}\"."
            if authors_found
            else f"[WARN] No se encontraron plantillas en \"{target}\"."
        )
        return AuthorCheckResult(True, message, authors_found)

    if not validation_enabled:
        return AuthorCheckResult(True, "[INFO] Validación de autores deshabilitada.", [])

    if target.suffix.lower() == ".thmx":
        return AuthorCheckResult(True, "[INFO] Validación de autor omitida para temas.", [])

    author, error = _extract_author(target)
    if error:
        return AuthorCheckResult(False, error, [], error=True)
    if not author:
        return AuthorCheckResult(False, f"[WARN] El archivo \"{target}\" no tiene autor asignado.", [])

    is_allowed = any(author.lower() == a.lower() for a in allowed)
    message = "[OK] Autor aprobado." if is_allowed else f"[BLOCKED] Autor no permitido para \"{target}\"."
    return AuthorCheckResult(is_allowed, message, [author])


def _normalize_allowed_authors(authors: Iterable[str]) -> list[str]:
    normalized: list[str] = []
    for author in authors:
        cleaned = author.strip()
        if cleaned:
            normalized.append(cleaned)
    return normalized


def _extract_author(template_path: Path) -> tuple[Optional[str], Optional[str]]:
    if not template_path.exists():
        return None, f"[ERROR] No se encontró la ruta: \"{template_path}\""

    try:
        with zipfile.ZipFile(template_path) as zipped:
            try:
                with zipped.open("docProps/core.xml") as core_file:
                    tree = ET.fromstring(core_file.read())
            except KeyError:
                return None, f"[WARN] No se pudo obtener el autor para \"{template_path.name}\" (core.xml ausente)."
    except Exception as exc:  # noqa: BLE001
        return None, f"[ERROR] {template_path.name}: {exc}"

    for candidate in ("{http://purl.org/dc/elements/1.1/}creator", "creator"):
        node = tree.find(candidate)
        if node is not None and node.text:
            return node.text.strip(), None
    return None, f"[WARN] \"{template_path.name}\" sin autor definido."
