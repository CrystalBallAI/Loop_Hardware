#!/usr/bin/env bash
# build_macos_nuitka.sh — HARDENED build: compile to native machine code with
# Nuitka (no .pyc bytecode to extract, unlike PyInstaller). This is the
# recommended SHIPPING compiler; PyInstaller (build_macos.sh) is the fast-iterate
# fallback. Nuitka compiles the whole dependency graph to C → slow (tens of
# minutes) but produces the real code protection.
#
#   ./build/build_macos_nuitka.sh [--tester ID] [--days N]
set -euo pipefail
cd "$(dirname "$0")/.."

TESTER="dev"; DAYS=14
while [ $# -gt 0 ]; do
  case "$1" in
    --tester) TESTER="$2"; shift 2;;
    --days) DAYS="$2"; shift 2;;
    *) echo "unknown arg: $1"; exit 2;;
  esac
done

APPNAME="CBMI Loop"
DIST="dist-nuitka"
APP="$DIST/app.app"                 # nuitka names it after the entry (app.py)
FINAL="$DIST/$APPNAME.app"
DMG="$DIST/CBMI-Loop-${TESTER}-nuitka.dmg"

echo "==> icon + vendor(encrypt) + stamp"
python3 build/make_icon.py
python3 vendor.py --encrypt-only
python3 build/inject_build_info.py --tester "$TESTER" --days "$DAYS"

echo "==> nuitka compile (this takes a while)"
rm -rf "$DIST"
python3 -m nuitka \
  --standalone \
  --macos-create-app-bundle \
  --macos-app-name="$APPNAME" \
  --macos-app-icon=build/assets/icon.icns \
  --macos-signed-app-name="ai.crystalball.cbmiloop" \
  --company-name="Crystalball AI" \
  --product-name="CBMI Loop" \
  --product-version="0.1.0" \
  --include-package=cbmi \
  --include-package=cbmi_pipelines \
  --include-package=server \
  --include-package=uvicorn \
  --include-module=multipart \
  --include-data-dir=webui=webui \
  --include-data-dir=runtime_data=runtime_data \
  --nofollow-import-to=tkinter \
  --nofollow-import-to=matplotlib \
  --nofollow-import-to=pytest \
  --nofollow-import-to=IPython \
  --nofollow-import-to=pyarrow \
  --nofollow-import-to=numba \
  --nofollow-import-to=llvmlite \
  --nofollow-import-to=mypy \
  --nofollow-import-to=scipy \
  --nofollow-import-to=setuptools \
  --assume-yes-for-downloads \
  --output-dir="$DIST" \
  app.py

python3 build/inject_build_info.py --reset

# nuitka outputs app.app -> rename to the product name
[ -d "$APP" ] && mv "$APP" "$FINAL"

echo "==> ad-hoc codesign"
codesign --force --deep --sign - "$FINAL"
codesign --verify --deep --strict "$FINAL" && echo "    codesign verify OK"

echo "==> DMG"
TMP="$(mktemp -d)"; cp -R "$FINAL" "$TMP/"; ln -s /Applications "$TMP/Applications"
hdiutil create -volname "$APPNAME" -srcfolder "$TMP" -ov -format UDZO "$DMG" >/dev/null
rm -rf "$TMP"
echo "Built: $DMG ($(du -h "$DMG" | cut -f1))"
