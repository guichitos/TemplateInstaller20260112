"""Elimina plantillas Normal de Word en una ruta fija."""
from __future__ import annotations

from pathlib import Path

try:
    from . import common
except ImportError:  # pragma: no cover - permite ejecución directa como script
    import sys

    sys.path.append(str(Path(__file__).resolve().parent))
    import common  # type: ignore[no-redef]


def delete_normal_templates() -> None:
    design_mode = common.DEFAULT_DESIGN_MODE
    common.refresh_design_log_flags(design_mode)
    if design_mode:
        common.configure_logging(design_mode)
        common.remove_normal_templates(design_mode, emit=print)
    else:
        common.remove_normal_templates(design_mode)


if __name__ == "__main__":
    delete_normal_templates()
