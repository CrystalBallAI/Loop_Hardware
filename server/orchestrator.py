"""
orchestrator.py — desktop edition.

Stages uploaded inputs into the writable workspace, runs each subsystem's
4-step pipeline via cbmi.runner (spawned children of THIS process — no system
python3 needed, works frozen), builds the hardware-page result, and persists
per-run history under app-data.

Dev CLI (stages the SOURCE repo's sample_data through the real staging path):
    python3 -m server.orchestrator --subsystem checkpoint --sample
    python3 -m server.orchestrator --subsystem all --sample
"""
from __future__ import annotations

import argparse
import json
import shutil
import threading
import time
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from cbmi import appdata, runner
from cbmi.resources import resource_root

from . import run_store
from .adapters import build_result
from .subsystems import SUBSYSTEMS, Subsystem, get

STEP_TIMEOUT = {
    "pipeline": 1800,
    "recommendations": 300,
    "excel": 300,
    "provenance": 120,
}
CRITICAL_STEPS = {"pipeline", "recommendations"}


# ---------------------------------------------------------------------------
# Input staging
# ---------------------------------------------------------------------------
def prepare_subsystem(sub: Subsystem) -> None:
    """Wipe + recreate input_data/<sub>/ so each run starts clean."""
    root = sub.input_root
    if root.exists():
        shutil.rmtree(root)
    root.mkdir(parents=True, exist_ok=True)


def destination_for(sub: Subsystem, field: str, filename: str
                    ) -> Optional[Tuple[Path, Optional[Path]]]:
    """Resolve where one uploaded file should land. Returns (dest, also_copy) or None.

    Field name contract:
      non per-point: "<key>__<input_id>"            e.g. drone__drone_images
      per-point:     "<key>__point<N>__<slot_id>"   e.g. gcp__point0__rinex
    """
    parts = field.split("__")
    safe_name = Path(filename).name or "file"

    if not sub.per_point and len(parts) == 2:
        input_id = parts[1]
        for slot in sub.slots:
            if slot.input_id != input_id:
                continue
            if slot.kind == "dir":
                dest = sub.input_root / slot.dest / safe_name
            else:
                dest = sub.input_root / slot.dest
            also = (sub.codebase / slot.also_copy_to) if slot.also_copy_to else None
            return dest, also
        return None

    if sub.per_point and len(parts) == 3 and parts[1].startswith("point"):
        try:
            idx = int(parts[1][len("point"):])
        except ValueError:
            return None
        slot_id = parts[2]
        for ps in sub.point_slots:
            if ps.slot_id != slot_id:
                continue
            folder = sub.input_root / f"{sub.point_folder_prefix}{idx + 1}"
            if ps.kind == "dir":
                dest = folder / safe_name
            else:
                dest = folder / ps.dest
            return dest, None
        return None

    return None


def validate_inputs(sub: Subsystem) -> List[str]:
    """Light pre-check: catch the obvious 'required input absent' cases."""
    errs: List[str] = []
    root = sub.input_root
    if not root.exists():
        return [f"No inputs uploaded for {sub.label}."]

    if not sub.per_point:
        for slot in sub.slots:
            if not slot.required:
                continue
            if slot.kind == "dir":
                d = root / slot.dest
                if not d.exists() or not any(d.iterdir()):
                    errs.append(f"{sub.label}: '{slot.input_id}' is required (no files).")
            else:
                if not (root / slot.dest).exists():
                    errs.append(f"{sub.label}: '{slot.input_id}' is required.")
    else:
        point_dirs = sorted(d for d in root.glob(f"{sub.point_folder_prefix}*") if d.is_dir())
        if not point_dirs:
            errs.append(f"{sub.label}: add at least one point with its files.")
        for d in point_dirs:
            for ps in sub.point_slots:
                if not ps.required:
                    continue
                if ps.kind == "dir":
                    if not any(d.iterdir()):
                        errs.append(f"{sub.label} [{d.name}]: '{ps.slot_id}' is required.")
                else:
                    if not (d / ps.dest).exists():
                        errs.append(f"{sub.label} [{d.name}]: '{ps.slot_id}' is required.")
    return errs


def write_run_paths(sub: Subsystem) -> str:
    """Write paths.run.json at the WORKSPACE root with inputs pointed at
    ../input_data/<key>/ via lexical relative paths (input_data is a sibling of
    the per-subsystem workspace roots, preserving the codebase-era contract)."""
    cfg = json.loads((sub.codebase / "paths.json").read_text(encoding="utf-8"))
    inputs = cfg.setdefault("inputs", {})
    for key, dest in sub.inputs_override.items():
        if dest in (".", ""):
            rel = f"../input_data/{sub.key}"
        else:
            rel = f"../input_data/{sub.key}/{dest}"
        inputs[key] = rel
    sub.run_paths_file.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
    return sub.run_paths_file.name


# ---------------------------------------------------------------------------
# Results + run history (persisted in app-data)
# ---------------------------------------------------------------------------
def save_result(sub: Subsystem, result: dict) -> None:
    appdata.save_json(appdata.results_dir() / f"{sub.key}.json", result)


def load_result(key: str) -> Optional[dict]:
    return appdata.load_json(appdata.results_dir() / f"{key}.json")


def _persist_run_history(run_id: str) -> None:
    """Snapshot a finished run: status + per-subsystem result/xlsx/provenance.
    Called by whichever subsystem thread finishes last (idempotent)."""
    run = run_store.get_run(run_id)
    if not run or run.get("status") != "done":
        return
    rdir = appdata.runs_dir() / run_id
    run["finishedAt"] = int(time.time())
    appdata.save_json(rdir / "run.json", run)
    for key, s in run["subsystems"].items():
        if s.get("status") != "done":
            continue
        sub = SUBSYSTEMS[key]
        sdir = rdir / key
        sdir.mkdir(parents=True, exist_ok=True)
        res = load_result(key)
        if res:
            appdata.save_json(sdir / "result.json", res)
        xlsx = sub.codebase / sub.xlsx_name
        if xlsx.exists():
            shutil.copyfile(xlsx, sdir / xlsx.name)
        prov = _provenance_file(sub)
        if prov:
            shutil.copyfile(prov, sdir / prov.name)


def _provenance_file(sub: Subsystem) -> Optional[Path]:
    paths = appdata.load_json(sub.codebase / "paths.json") or {}
    spec_rel = paths.get("spec_file", "")
    spec_dir = (sub.codebase / spec_rel).parent if spec_rel else sub.codebase
    candidates = sorted(spec_dir.glob("*provenance*.html"))
    return candidates[-1] if candidates else None


def list_runs() -> List[dict]:
    out = []
    root = appdata.runs_dir()
    if not root.exists():
        return out
    for d in sorted(root.iterdir(), reverse=True):
        run = appdata.load_json(d / "run.json")
        if not run:
            continue
        out.append({
            "runId": run.get("runId", d.name),
            "finishedAt": run.get("finishedAt"),
            "subsystems": {
                k: {f: s.get(f) for f in ("status", "score", "grade", "decision")}
                for k, s in (run.get("subsystems") or {}).items()
            },
        })
    return out


def get_run_history(run_id: str) -> Optional[dict]:
    return appdata.load_json(appdata.runs_dir() / run_id / "run.json")


def activate_run(run_id: str) -> List[str]:
    """Make a historical run's results the ones the hardware pages show."""
    activated = []
    rdir = appdata.runs_dir() / run_id
    for key in SUBSYSTEMS:
        res = appdata.load_json(rdir / key / "result.json")
        if res:
            appdata.save_json(appdata.results_dir() / f"{key}.json", res)
            activated.append(key)
    return activated


# ---------------------------------------------------------------------------
# Run a subsystem (one thread per subsystem; steps are spawned child processes)
# ---------------------------------------------------------------------------
def run_subsystem(run_id: str, sub: Subsystem) -> None:
    run_store.set_subsystem_status(run_id, sub.key, "running")

    run_store.set_step(run_id, sub.key, "validate", "running")
    errs = validate_inputs(sub)
    paths_arg = None if errs else write_run_paths(sub)
    if errs:
        run_store.set_step(run_id, sub.key, "validate", "failed")
        run_store.set_subsystem_status(run_id, sub.key, "failed")
        run_store.set_result(run_id, sub.key, error="; ".join(errs[:4]))
        run_store.finalize(run_id)
        _persist_run_history(run_id)
        return
    run_store.set_step(run_id, sub.key, "validate", "done")

    failed = False
    for step_id in ("pipeline", "recommendations", "excel", "provenance"):
        run_store.set_step(run_id, sub.key, step_id, "running")
        rc, tail = runner.run_step(
            sub.key, step_id,
            None if step_id == "provenance" else paths_arg,
            cwd=sub.codebase,
            log_file=appdata.logs_dir() / f"{sub.key}-{step_id}.log",
            timeout=STEP_TIMEOUT.get(step_id, 300),
        )
        run_store.append_log(run_id, sub.key, f"[{step_id}] rc={rc}")
        if rc != 0:
            run_store.set_step(run_id, sub.key, step_id, "failed")
            run_store.append_log(run_id, sub.key, tail.strip()[-400:])
            if step_id in CRITICAL_STEPS:
                run_store.set_subsystem_status(run_id, sub.key, "failed")
                run_store.set_result(run_id, sub.key,
                                     error=f"{step_id} failed: {tail.strip()[-200:]}")
                failed = True
                break
            continue
        run_store.set_step(run_id, sub.key, step_id, "done")
        if step_id == "pipeline":
            _publish_partial_score(run_id, sub)

    try:
        result = build_result(sub)
        save_result(sub, result)
        run_store.set_result(run_id, sub.key,
                             score=result.get("overallScore"),
                             grade=result.get("grade"),
                             decision=result.get("decisionLabel"))
        if not failed:
            run_store.set_subsystem_status(run_id, sub.key, "done")
    except Exception as exc:  # noqa: BLE001
        run_store.append_log(run_id, sub.key, f"adapter error: {exc}")
        if not failed:
            run_store.set_subsystem_status(run_id, sub.key, "failed")
            run_store.set_result(run_id, sub.key, error=f"result build failed: {exc}")

    run_store.finalize(run_id)
    _persist_run_history(run_id)


def _publish_partial_score(run_id: str, sub: Subsystem) -> None:
    try:
        obj = json.loads((sub.outputs_dir / Path(sub.score_file).name)
                         .read_text(encoding="utf-8"))
        data = obj.get("data", obj)
        score = data.get(sub.score_key)
        if isinstance(score, (int, float)):
            run_store.set_result(run_id, sub.key, score=round(score))
    except Exception:  # noqa: BLE001
        pass


def start_run(subsystem_keys: List[str]) -> str:
    run_id = run_store.create_run(subsystem_keys)
    for key in subsystem_keys:
        t = threading.Thread(target=run_subsystem,
                             args=(run_id, SUBSYSTEMS[key]), daemon=True)
        t.start()
    return run_id


# ---------------------------------------------------------------------------
# Dev-only sample staging: copy the SOURCE repo's sample_data through the real
# staging layout (so a sample run exercises the exact production code path).
# ---------------------------------------------------------------------------
def stage_sample(sub: Subsystem) -> None:
    src_root = resource_root().parent / sub.source_codebase
    src_paths = json.loads((src_root / "paths.json").read_text(encoding="utf-8"))
    prepare_subsystem(sub)

    if sub.per_point:
        points_root = src_root / src_paths["inputs"]["points_root"]
        glob = src_paths["inputs"].get("point_folder_glob",
                                       f"{sub.point_folder_prefix}*")
        for d in sorted(points_root.glob(glob)):
            if d.is_dir():
                shutil.copytree(d, sub.input_root / d.name)
        return

    for pkey, dest in sub.inputs_override.items():
        src_rel = src_paths["inputs"].get(pkey)
        if not src_rel:
            continue
        src = src_root / src_rel
        target = sub.input_root / dest
        if src.is_dir():
            shutil.copytree(src, target)
        elif src.is_file():
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(src, target)
    # base station: the hardware override lives at sample_data/hardware.json
    hw = src_root / "sample_data" / "hardware.json"
    if hw.exists():
        mirror = sub.codebase / "sample_data" / "hardware.json"
        mirror.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(hw, mirror)


def _cli() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--subsystem", default="all",
                    help="drone|basestation|gcp|checkpoint|all")
    ap.add_argument("--sample", action="store_true",
                    help="stage the source repo's sample_data as inputs first")
    args = ap.parse_args()

    appdata.materialize()
    keys = list(SUBSYSTEMS) if args.subsystem == "all" else [get(args.subsystem).key]
    if args.sample:
        for key in keys:
            stage_sample(SUBSYSTEMS[key])
            print(f"staged sample inputs for {key}")

    run_id = run_store.create_run(keys)
    threads = []
    for key in keys:
        t = threading.Thread(target=run_subsystem, args=(run_id, SUBSYSTEMS[key]))
        t.start()
        threads.append(t)
    for t in threads:
        t.join()

    run = run_store.get_run(run_id)
    print("\n==== RUN SUMMARY ====")
    for key, s in run["subsystems"].items():
        steps = " ".join(f"{st['id']}:{st['status']}" for st in s["steps"])
        print(f"\n{key:12s} status={s['status']} score={s['score']} "
              f"grade={s['grade']} decision={s['decision']}")
        print(f"  steps: {steps}")
        if s["error"]:
            print(f"  error: {s['error']}")
        res = load_result(key)
        if res:
            print(f"  result: overall={res['overallScore']} tier={res.get('tier')!r} "
                  f"blocks={[(b['id'], b['score']) for b in res['bbs']]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(_cli())
