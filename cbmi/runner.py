"""Run one pipeline step in a child process of THIS interpreter/binary.

Replaces the dev-era `subprocess([python3, scripts/xxx.py, paths.run.json])`:
a frozen app has no python3, and its pipeline code exists only as compiled
modules, not .py files. So the parent spawns a multiprocessing child (spawn
context — the only one that behaves identically frozen and unfrozen, and the
Windows default) which imports the vendored package and calls its main().

Why a process per step (not a thread):
  * os.chdir / env are process-global — parallel subsystems would race;
  * a native-extension crash (georinex/numpy) must not take down the app;
  * timeouts stay enforceable (terminate the child).

Child stdout/stderr go to logs/<key>-<step>.log; the parent returns the tail
for run-status display, same contract as the old subprocess runner.
"""
from __future__ import annotations

import multiprocessing as mp
import threading
from pathlib import Path

STEP_MODULES = {
    "pipeline": "run_pipeline",
    "recommendations": "compute_recommendations",
    "excel": "export_results_xlsx",
    "provenance": "build_provenance_html",
}

_ctx = mp.get_context("spawn")
_active_lock = threading.Lock()
_active: set = set()


def _step_child(key: str, step: str, paths_arg, cwd: str, log_file: str) -> None:
    """Child-process entry. Must stay importable at module top level (spawn
    re-imports this module in the child)."""
    import importlib
    import os
    import sys
    import traceback

    rc = 1
    try:
        log = open(log_file, "w", buffering=1, encoding="utf-8", errors="replace")
        sys.stdout = sys.stderr = log
        os.chdir(cwd)
        os.environ["CBMI_PIPELINE_ROOT"] = cwd

        # transparent decryption of the encrypted spec/library files this step reads
        from cbmi import crypt
        crypt.install_read_hook()

        script = STEP_MODULES[step]
        mod = importlib.import_module(f"cbmi_pipelines.{key}.scripts.{script}")
        argv = [] if paths_arg is None else [str(paths_arg)]
        sys.argv = [script] + argv

        # replay the script's own __main__ call (vendor-generated table), so
        # each script gets argv exactly the way its source expects it
        from cbmi_pipelines.entrypoints import ENTRY
        call = ENTRY[f"{key}.{script}"]
        try:
            rc = eval(call, {"main": mod.main, "sys": sys})  # noqa: S307
        except SystemExit as exc:
            rc = exc.code if isinstance(exc.code, int) else (0 if exc.code is None else 1)
        rc = int(rc or 0)
    except Exception:  # noqa: BLE001
        traceback.print_exc()
        rc = 1
    finally:
        try:
            sys.stdout.flush()
        except Exception:  # noqa: BLE001
            pass
        os._exit(rc)


def run_step(key: str, step: str, paths_arg, cwd: Path, log_file: Path,
             timeout: int) -> tuple:
    """Blocking. Returns (rc, output_tail)."""
    log_file.parent.mkdir(parents=True, exist_ok=True)
    proc = _ctx.Process(
        target=_step_child,
        args=(key, step, paths_arg, str(cwd), str(log_file)),
        daemon=True,
    )
    proc.start()
    with _active_lock:
        _active.add(proc)
    try:
        proc.join(timeout)
        if proc.is_alive():
            proc.terminate()
            proc.join(10)
            if proc.is_alive():
                proc.kill()
                proc.join(5)
            return 124, f"{step} timed out after {timeout}s"
        rc = proc.exitcode if proc.exitcode is not None else 1
    finally:
        with _active_lock:
            _active.discard(proc)

    tail = ""
    try:
        text = log_file.read_text(encoding="utf-8", errors="replace")
        tail = text[-3000:]
    except OSError:
        pass
    return rc, tail


def terminate_all() -> None:
    """App shutdown: kill any still-running pipeline steps."""
    with _active_lock:
        procs = list(_active)
    for p in procs:
        try:
            if p.is_alive():
                p.terminate()
        except Exception:  # noqa: BLE001
            pass
