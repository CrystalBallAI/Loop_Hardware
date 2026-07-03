"""Per-user writable state. The install dir (Program Files / /Applications) is
read-only, so everything a run writes lives under the user's app-data dir:

  <appdata>/CBMI Loop/
    workspace/<key>/           materialized pipeline root: paths.json, spec,
                               recommendations lib, cache/, outputs/  (per run)
    workspace/input_data/<key> staged uploads — sibling of the pipeline roots
                               because paths.run.json uses lexical
                               ../input_data/<key>/ paths (the pipelines call
                               .relative_to(root), which is lexical)
    results/                   latest hardware-page result per subsystem
    runs/                      per-run history snapshots
    logs/                      app + per-step logs
"""
from __future__ import annotations

import json
import os
import platform
import shutil
from pathlib import Path

from . import APP_NAME, APP_VERSION
from .resources import runtime_data_dir


def app_data_dir() -> Path:
    sysname = platform.system()
    if sysname == "Darwin":
        base = Path.home() / "Library" / "Application Support"
    elif sysname == "Windows":
        base = Path(os.environ.get("LOCALAPPDATA")
                    or (Path.home() / "AppData" / "Local"))
    else:
        base = Path(os.environ.get("XDG_DATA_HOME")
                    or (Path.home() / ".local" / "share"))
    return base / APP_NAME


def workspace_dir() -> Path:
    return app_data_dir() / "workspace"


def pipeline_root(key: str) -> Path:
    return workspace_dir() / key


def input_data_root() -> Path:
    return workspace_dir() / "input_data"


def results_dir() -> Path:
    return app_data_dir() / "results"


def runs_dir() -> Path:
    return app_data_dir() / "runs"


def logs_dir() -> Path:
    return app_data_dir() / "logs"


_STAMP = "workspace.version"


def _install_crypt() -> None:
    """Parent process reads the spec too (adapters.build_result) — install the
    same transparent-decryption hook here. Idempotent."""
    try:
        from . import crypt
        crypt.install_read_hook()
    except Exception:  # noqa: BLE001 — dev without cryptography still runs (plaintext)
        pass


def materialize(force: bool = False) -> None:
    """Copy bundled runtime_data/<key>/ into the writable workspace. Re-done
    whenever the app version changes (or force=True) so spec/library updates in
    a new build reach the workspace; input_data/, results/, runs/ are kept."""
    src = runtime_data_dir()
    ws = workspace_dir()
    _install_crypt()
    stamp = app_data_dir() / _STAMP
    current = stamp.read_text(encoding="utf-8").strip() if stamp.exists() else ""
    if current == APP_VERSION and not force and ws.exists():
        return

    for key_dir in sorted(p for p in src.iterdir() if p.is_dir()):
        dest = ws / key_dir.name
        if dest.exists():
            shutil.rmtree(dest)
        shutil.copytree(key_dir, dest)

    for d in (input_data_root(), results_dir(), runs_dir(), logs_dir()):
        d.mkdir(parents=True, exist_ok=True)
    stamp.parent.mkdir(parents=True, exist_ok=True)
    stamp.write_text(APP_VERSION, encoding="utf-8")


def save_json(path: Path, obj) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, indent=2), encoding="utf-8")


def load_json(path: Path):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
