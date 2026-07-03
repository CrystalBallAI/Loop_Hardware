#!/usr/bin/env bash
# build_macos.sh — produce a signed-ad-hoc .app + .dmg for CBMI Loop.
#
#   ./build/build_macos.sh [--tester ID] [--days N]
#
# Unsigned for beta (user's choice): we ad-hoc codesign only, which is the
# MINIMUM Apple Silicon needs to launch at all. Not notarized — testers use the
# "Open Anyway" flow (see TESTER_GUIDE.md). Run from the Desktop_App dir.
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
DIST="dist"
APP="$DIST/$APPNAME.app"
DMG="$DIST/CBMI-Loop-${TESTER}.dmg"

echo "==> [1/6] icon"
python3 build/make_icon.py

echo "==> [2/6] vendor + ENCRYPT scoring IP"
python3 vendor.py --encrypt

echo "==> [3/6] stamp build identity (tester=$TESTER, expiry=${DAYS}d)"
python3 build/inject_build_info.py --tester "$TESTER" --days "$DAYS"

echo "==> [4/6] freeze (.app)"
rm -rf build/pyi "$APP"
python3 -m PyInstaller --noconfirm --clean --workpath build/pyi --distpath "$DIST" cbmi_loop.spec

# reset the checked-in build_info so the repo never keeps a tester stamp
python3 build/inject_build_info.py --reset

echo "==> [5/6] ad-hoc codesign (required for arm64 launch)"
codesign --force --deep --sign - "$APP"
codesign --verify --deep --strict "$APP" && echo "    codesign verify OK"

echo "==> [6/6] DMG"
rm -f "$DMG"
TMP="$(mktemp -d)"
cp -R "$APP" "$TMP/"
ln -s /Applications "$TMP/Applications"
hdiutil create -volname "$APPNAME" -srcfolder "$TMP" -ov -format UDZO "$DMG" >/dev/null
rm -rf "$TMP"

echo ""
echo "Built:"
echo "  app: $APP"
echo "  dmg: $DMG  ($(du -h "$DMG" | cut -f1))"
