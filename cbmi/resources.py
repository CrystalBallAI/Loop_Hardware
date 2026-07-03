"""Locate read-only bundled resources (webui/, runtime_data/) in both dev and
frozen layouts.

Dev:     Desktop_App/{webui,runtime_data}       (this file lives in Desktop_App/cbmi/)
Frozen:  <dist dir>/{webui,runtime_data}        (folders shipped beside the binary;
         PyInstaller onedir and Nuitka standalone both put support files next to
         the executable)
"""
from __future__ import annotations

import sys
from pathlib import Path


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False)) or "__compiled__" in globals()


def resource_root() -> Path:
    if is_frozen():
        # PyInstaller onedir: bundled data lives in _internal (sys._MEIPASS);
        # Nuitka standalone: data folders sit beside the executable.
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            return Path(meipass)
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def webui_dir() -> Path:
    return resource_root() / "webui"


def runtime_data_dir() -> Path:
    return resource_root() / "runtime_data"
