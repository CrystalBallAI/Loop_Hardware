"""
In-memory run status store.

A "run" is one click of Validate & Run Scoring. It contains one entry per
subsystem, each tracking the 5 pipeline steps. The orchestrator updates this
as subprocesses progress; the frontend polls GET /api/run/{id}.

Thread-safe: subsystems run in parallel threads, all writing here.
"""
from __future__ import annotations

import threading
from typing import Any, Dict, List, Optional

from .subsystems import STEP_ORDER, STEP_LABELS, SUBSYSTEMS

_lock = threading.RLock()
_runs: Dict[str, Dict[str, Any]] = {}
_counter = {"n": 0}


def _new_steps() -> List[Dict[str, str]]:
    return [{"id": s, "label": STEP_LABELS[s], "status": "pending"} for s in STEP_ORDER]


def create_run(subsystem_keys: List[str]) -> str:
    with _lock:
        _counter["n"] += 1
        # timestamped so persisted run history never collides across app restarts
        import time
        stamp = time.strftime("%Y%m%d-%H%M%S")
        run_id = f"run-{stamp}-{_counter['n']:03d}"
        _runs[run_id] = {
            "runId": run_id,
            "status": "running",
            "subsystems": {
                k: {
                    "key": k,
                    "label": SUBSYSTEMS[k].label,
                    "module": SUBSYSTEMS[k].module,
                    "status": "queued",          # queued|running|done|failed|skipped
                    "steps": _new_steps(),
                    "score": None,
                    "grade": None,
                    "decision": None,
                    "error": None,
                    "log": [],
                }
                for k in subsystem_keys
            },
        }
        return run_id


def _sub(run_id: str, key: str) -> Optional[Dict[str, Any]]:
    run = _runs.get(run_id)
    if not run:
        return None
    return run["subsystems"].get(key)


def set_subsystem_status(run_id: str, key: str, status: str) -> None:
    with _lock:
        s = _sub(run_id, key)
        if s:
            s["status"] = status


def set_step(run_id: str, key: str, step_id: str, status: str) -> None:
    with _lock:
        s = _sub(run_id, key)
        if not s:
            return
        for st in s["steps"]:
            if st["id"] == step_id:
                st["status"] = status
                break


def append_log(run_id: str, key: str, line: str) -> None:
    with _lock:
        s = _sub(run_id, key)
        if s is not None and line:
            s["log"].append(line)
            # keep the tail bounded
            if len(s["log"]) > 60:
                s["log"] = s["log"][-60:]


def set_result(run_id: str, key: str, score=None, grade=None, decision=None,
               error=None) -> None:
    with _lock:
        s = _sub(run_id, key)
        if not s:
            return
        if score is not None:
            s["score"] = score
        if grade is not None:
            s["grade"] = grade
        if decision is not None:
            s["decision"] = decision
        if error is not None:
            s["error"] = error


def finalize(run_id: str) -> None:
    """Mark the overall run done once every subsystem is terminal."""
    with _lock:
        run = _runs.get(run_id)
        if not run:
            return
        terminal = {"done", "failed", "skipped"}
        if all(s["status"] in terminal for s in run["subsystems"].values()):
            run["status"] = "done"


def get_run(run_id: str) -> Optional[Dict[str, Any]]:
    with _lock:
        run = _runs.get(run_id)
        if run is None:
            return None
        # shallow-ish copy is fine for JSON serialization under the lock
        import copy
        return copy.deepcopy(run)
