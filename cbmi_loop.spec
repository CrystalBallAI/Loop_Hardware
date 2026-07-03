# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for CBMI Loop (macOS .app / Windows onedir).

Bundles the ENCRYPTED runtime_data (run `python3 vendor.py --encrypt` first) so
the scoring IP ships as ciphertext. Trims the biggest non-dependencies that a
kitchen-sink (e.g. Anaconda) environment drags in; a clean venv build is smaller
still. Build via build/build_macos.sh or `pyinstaller cbmi_loop.spec`.
"""
import sys
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

datas = [
    ("webui", "webui"),
    ("runtime_data", "runtime_data"),
]
# georinex/xarray ship data files + version markers loaded at runtime
for pkg in ("georinex", "xarray"):
    try:
        datas += collect_data_files(pkg)
    except Exception:
        pass

hiddenimports = []
for pkg in ("cbmi_pipelines", "cbmi", "server", "uvicorn"):
    hiddenimports += collect_submodules(pkg)
# multipart + cryptography backends PyInstaller's static analysis can miss
hiddenimports += ["multipart", "cryptography.hazmat.backends.openssl"]

excludes = [
    "tkinter", "matplotlib", "scipy.spatial.cKDTree",
    "botocore", "boto3", "panel", "bokeh", "IPython", "jupyter",
    "notebook", "pytest", "sphinx", "PyQt6", "PySide2", "PySide6",
    "test", "tests", "setuptools._vendor",
]

block_cipher = None

a = Analysis(
    ["app.py"],
    pathex=["."],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=excludes,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name="CBMI-Loop",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,          # windowed — no terminal window
    disable_windowed_traceback=False,
    target_arch=None,       # host arch (arm64 on Apple Silicon)
    codesign_identity=None,
    entitlements_file=None,
    icon="build/assets/icon.ico" if sys.platform.startswith("win") else "build/assets/icon.icns",
)
coll = COLLECT(
    exe, a.binaries, a.datas,
    strip=False, upx=False, upx_exclude=[], name="CBMI-Loop",
)

if sys.platform == "darwin":
    app = BUNDLE(
        coll,
        name="CBMI Loop.app",
        icon="build/assets/icon.icns",
        bundle_identifier="ai.crystalball.cbmiloop",
        version="0.1.0",
        info_plist={
            "CFBundleName": "CBMI Loop",
            "CFBundleDisplayName": "CBMI Loop",
            "CFBundleShortVersionString": "0.1.0",
            "NSHighResolutionCapable": True,
            "LSMinimumSystemVersion": "11.0",
            # server binds 127.0.0.1 only; no ATS exceptions needed for remote,
            # but the embedded webview loads http://127.0.0.1 → allow local loopback
            "NSAppTransportSecurity": {"NSAllowsLocalNetworking": True},
        },
    )
