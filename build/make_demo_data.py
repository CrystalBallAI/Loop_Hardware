#!/usr/bin/env python3
"""
make_demo_data.py — assemble a small demo dataset zip that testers can load to
try the app in minutes, WITHOUT shipping it inside the installer (keeps the app
small; the demo is optional and shared separately via Drive).

Pulls the lightest real sample inputs from the source *_CodeBase sample_data
(Check Point + Control Point — both small; Drone's 3 GB imagery is excluded).

Usage:  python3 build/make_demo_data.py
Output: demo_data.zip  (+ a README explaining which files map to which slot)
"""
from __future__ import annotations

import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parents[1]
REPO = HERE.parent
OUT = HERE / "demo_data.zip"

# (source codebase, sample subpath, arcname-in-zip). Only small subsystems.
SETS = {
    "CheckPoint": ("CheckPoint_CodeBase", "sample_data", "check_point"),
    "ControlPoint": ("GCP_CodeBase", "sample_data", "control_point"),
    "BaseStation": ("BaseStation_CodeBase", "sample_data", "base_station"),
}
MAX_FILE_MB = 40           # skip anything larger (keeps the demo lightweight)

README = """CBMI Loop — Demo Data
=====================
A small real dataset for trying the app. In CBMI Loop, open the Load page,
pick the matching tab, and upload the files from the matching folder here:

  check_point/   -> Check Point tab  (per point: cp_rtk_export.csv + cp_user_input.json)
  control_point/ -> Control Point tab (per point: RINEX obs+nav + user_input.json)
  base_station/  -> Base tab          (base_rinex/ + operator_log + user_input)

Drone is omitted here — its imagery set is multiple GB. Ask us for a drone demo
separately if you want to exercise that pipeline.
"""


def main() -> int:
    if OUT.exists():
        OUT.unlink()
    n = 0
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("README.txt", README)
        for label, (cb, sub, arc) in SETS.items():
            root = REPO / cb / sub
            if not root.exists():
                print(f"  skip {label}: {root} missing")
                continue
            for p in sorted(root.rglob("*")):
                if not p.is_file():
                    continue
                if p.stat().st_size > MAX_FILE_MB * 1024 * 1024:
                    continue
                if p.name == ".DS_Store":
                    continue
                z.write(p, f"{arc}/{p.relative_to(root)}")
                n += 1
            print(f"  added {label}")
    print(f"wrote {OUT.name} ({n} files, {OUT.stat().st_size // 1024} KB)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
