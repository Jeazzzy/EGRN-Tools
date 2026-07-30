"""Prepare pyogrio's native runtime before application modules are imported."""

from __future__ import annotations

import os
import sys
from pathlib import Path


# os.add_dll_directory() removes the directory when its handle is closed, so
# these handles must stay alive for the full process lifetime.
_DLL_DIRECTORY_HANDLES = []


def _prepend_env_path(variable: str, directory: Path) -> None:
    current = os.environ.get(variable, "")
    value = str(directory)
    os.environ[variable] = value if not current else value + os.pathsep + current


if sys.platform == "win32" and getattr(sys, "frozen", False):
    bundle_dir = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    libs_dir = bundle_dir / "pyogrio.libs"

    if libs_dir.is_dir():
        _DLL_DIRECTORY_HANDLES.append(os.add_dll_directory(str(libs_dir)))
        _prepend_env_path("PATH", libs_dir)

    gdal_data = bundle_dir / "pyogrio" / "gdal_data"
    if gdal_data.is_dir():
        os.environ["GDAL_DATA"] = str(gdal_data)

    proj_data = bundle_dir / "pyogrio" / "proj_data"
    if proj_data.is_dir():
        os.environ["PROJ_DATA"] = str(proj_data)
        os.environ["PROJ_LIB"] = str(proj_data)
