"""Per-build identity: watermark + expiry.

The build pipeline (Phase 3) rewrites the two constants below per tester before
compiling, so they end up baked into native code, and the packaged app enforces
them. In dev builds both are inert.

Clock-rollback guard: we persist the newest timestamp ever seen; if the system
clock is behind it by more than a day, the expiry check refuses to run.
"""
from __future__ import annotations

import time
from pathlib import Path

# --- injected at build time (see build/inject_build_info.py, Phase 3) -------
TESTER_ID = "dev"
EXPIRY_EPOCH = 0
# ---------------------------------------------------------------------------

_ROLLBACK_SLACK = 86400   # 1 day


def check(state_dir: Path) -> tuple[bool, str]:
    """Returns (ok, message). message is a user-facing expiry note."""
    now = int(time.time())

    marker = state_dir / "last_seen"
    try:
        last_seen = int(marker.read_text(encoding="utf-8").strip() or "0")
    except (OSError, ValueError):
        last_seen = 0
    if now + _ROLLBACK_SLACK < last_seen:
        return False, ("System clock appears to have been set backwards. "
                       "Fix the date/time and relaunch.")
    if now > last_seen:
        try:
            state_dir.mkdir(parents=True, exist_ok=True)
            marker.write_text(str(now), encoding="utf-8")
        except OSError:
            pass

    if not EXPIRY_EPOCH:
        return True, ""
    if now > EXPIRY_EPOCH:
        return False, ("This beta build has expired. "
                       "Contact bhanu@crystalball.ai for a new build.")
    days = max(1, round((EXPIRY_EPOCH - now) / 86400))
    return True, f"Beta build · expires in {days} day{'s' if days != 1 else ''}"
