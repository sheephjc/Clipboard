from __future__ import annotations

from pathlib import Path
import sys


def app_bundle_dir(executable: Path | None = None) -> Path | None:
    executable = (executable or Path(sys.executable)).resolve()
    for parent in (executable, *executable.parents):
        if parent.suffix == ".app":
            return parent
    return None


def portable_data_dir(project_root: Path, data_name: str = "data") -> Path:
    if getattr(sys, "frozen", False):
        bundle_dir = app_bundle_dir()
        if bundle_dir is not None:
            return bundle_dir.parent / data_name
        return Path(sys.executable).resolve().parent / data_name
    return project_root.resolve() / data_name

