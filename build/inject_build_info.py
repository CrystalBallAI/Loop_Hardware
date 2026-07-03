#!/usr/bin/env python3
"""
inject_build_info.py — stamp a per-tester identity + expiry into cbmi/build_info.py
just before a release build. CI calls this with the tester id; the compiled
binary then enforces the expiry (see cbmi/build_info.check).

Usage:
    python3 build/inject_build_info.py --tester "acme-pilot" --days 14
    python3 build/inject_build_info.py --reset          # back to dev defaults
"""
from __future__ import annotations

import argparse
import re
import time
from pathlib import Path

BUILD_INFO = Path(__file__).resolve().parents[1] / "cbmi" / "build_info.py"


def _set(text: str, name: str, value: str) -> str:
    pat = re.compile(rf"^{name} = .*$", re.MULTILINE)
    if not pat.search(text):
        raise SystemExit(f"inject: constant {name} not found in build_info.py")
    return pat.sub(f"{name} = {value}", text, count=1)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--tester", default="dev")
    ap.add_argument("--days", type=int, default=14,
                    help="days until the beta build expires (0 = never)")
    ap.add_argument("--expiry-epoch", type=int, default=None,
                    help="explicit expiry unix time (overrides --days; for reproducible CI)")
    ap.add_argument("--reset", action="store_true")
    args = ap.parse_args()

    text = BUILD_INFO.read_text(encoding="utf-8")
    if args.reset:
        text = _set(text, "TESTER_ID", '"dev"')
        text = _set(text, "EXPIRY_EPOCH", "0")
    else:
        if args.expiry_epoch is not None:
            expiry = args.expiry_epoch
        elif args.days > 0:
            expiry = int(time.time()) + args.days * 86400
        else:
            expiry = 0
        text = _set(text, "TESTER_ID", repr(args.tester))
        text = _set(text, "EXPIRY_EPOCH", str(expiry))

    BUILD_INFO.write_text(text, encoding="utf-8")
    exp = "never" if args.reset or (not args.expiry_epoch and args.days <= 0) \
        else time.strftime("%Y-%m-%d", time.localtime(
            args.expiry_epoch if args.expiry_epoch else int(time.time()) + args.days * 86400))
    print(f"build_info: tester={'dev' if args.reset else args.tester} expiry={exp}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
