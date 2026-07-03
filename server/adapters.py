"""
adapters.py — transform a subsystem's pipeline outputs into the JSON shape the
hardware pages consume (the DRONE_BBS-style structure the SPA already renders).

Sources (all already produced by the existing pipelines, nothing rewritten):
  * outputs/06_<sub>_score.json   -> apex score, per-block scores, global gate
  * outputs/07_recommendations.json -> per-indicator score/band/impact/actions,
                                       decision, tier_interpretation, flags
  * <spec_file> (from paths.json) -> full per-indicator band ladders (thresholds)

The result is intentionally built from the pipeline's REAL block model (e.g. drone
has 3 scoring blocks img/mis/gnss; calibration is folded into img), not the SPA's
static 4-block mock — so the page shows real structure.
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from .subsystems import Subsystem

# Short ids used by the SPA hardware ribbons (best-effort; falls back to a slug).
BB_SHORT = {
    "BB_IMG_CAPTURE": "img",
    "BB_MISSION_EXEC": "mis",
    "BB_ROVER_GNSS": "gnss",
    "BB_CAL_CONF": "cal",
}

DECISION_LABEL = {
    "good_to_go": "GOOD TO GO",
    "review_recommended": "REVIEW",
    "resurvey_recommended": "RESURVEY",
    "unable_to_assess": "UNABLE TO ASSESS",
}
DECISION_VERDICT = {           # maps to the SPA's good/review/resurvey colour states
    "good_to_go": "good",
    "review_recommended": "review",
    "resurvey_recommended": "resurvey",
    "unable_to_assess": "review",
}


def _load(path: Path) -> Optional[Any]:
    try:
        with path.open("r", encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return None


def _envelope_data(obj: Any) -> Dict[str, Any]:
    """Pipeline files wrap payload in {data:{...}}; recommendations is flat."""
    if isinstance(obj, dict) and "data" in obj and isinstance(obj["data"], dict):
        return obj["data"]
    return obj if isinstance(obj, dict) else {}


def _tier_band(score_value: int) -> str:
    if score_value >= 95:
        return "Excellent"
    if score_value >= 85:
        return "Strong"
    if score_value >= 70:
        return "Acceptable"
    if score_value >= 50:
        return "Marginal"
    return "Critical"


def _short_id(bb_id: str) -> str:
    if bb_id in BB_SHORT:
        return BB_SHORT[bb_id]
    # keep distinct (e.g. BB_CP_COMPLETE -> cp_complete, not 'cp')
    return bb_id.lower().replace("bb_", "") or bb_id.lower()


def _thresholds_by_indicator(spec: Dict[str, Any]) -> Dict[str, List[Dict[str, Any]]]:
    out: Dict[str, List[Dict[str, Any]]] = {}
    for t in (spec.get("thresholds") or []):
        out.setdefault(t.get("indicator_id"), []).append(t)
    for k in out:
        out[k].sort(key=lambda t: t.get("band_order", 0))
    return out


def _build_grades(thresholds: List[Dict[str, Any]], score: Optional[float]) -> List[Dict[str, Any]]:
    grades = []
    for t in thresholds:
        sv = t.get("score_value")
        grades.append({
            "l": _tier_band(sv if isinstance(sv, (int, float)) else 0),
            "r": t.get("condition_text") or t.get("condition_expression") or "",
            "s": sv,
            "current": False,
            "flag": t.get("flag_raised"),
        })
    # mark the band whose score_value matches the indicator score (else nearest)
    if grades and score is not None:  # noqa: PLR2004
        exact = [g for g in grades if g["s"] == score]
        if exact:
            exact[0]["current"] = True
        else:
            nearest = min(grades, key=lambda g: abs((g["s"] or 0) - score))
            nearest["current"] = True
    return grades


def _indicator_rec(ind: Dict[str, Any]) -> str:
    actions = ind.get("actions")
    impact = ind.get("impact")
    parts = []
    if impact:
        parts.append(impact)
    if actions:
        if isinstance(actions, list):
            parts.append(" ".join(str(a) for a in actions))
        else:
            parts.append(str(actions))
    if not parts:
        return ind.get("verified_statement") or ""
    return " ".join(parts)


def _indicator_alert(ind: Dict[str, Any]) -> Optional[str]:
    if ind.get("hard_gate_fired"):
        spec = ind.get("gate_action_spec")
        return f"Hard gate fired: {spec}" if spec else "Hard gate fired."
    mb = ind.get("matched_band") or {}
    if mb.get("flag"):
        return f"Flag: {mb['flag']}"
    return None


def _block_meta(score06: Dict[str, Any], agg05: Dict[str, Any],
                disp: Dict[str, str]) -> List[Dict[str, Any]]:
    """Per-block {bb_id, name, score, weight, gate, notes} from either
    06.block_contributions{} (drone) or 06.contributions[] (base/gcp/checkpoint).
    `disp` maps block_id -> display_name (from the spec's building_blocks)."""
    bc = score06.get("block_contributions")
    if isinstance(bc, dict) and bc:
        return [{
            "bb_id": k, "name": v.get("block_display_name") or disp.get(k) or k,
            "score": v.get("block_score"), "weight": v.get("weight_in_ppk"),
            "gate": bool(v.get("block_gate_triggered")), "notes": v.get("notes"),
        } for k, v in bc.items()]
    out = []
    for c in (score06.get("contributions") or []):
        bb = c.get("block_id")
        meta = (agg05 or {}).get(bb, {})
        score = c.get("block_aggregate_score")
        if score is None:
            score = c.get("block_score")
        out.append({
            "bb_id": bb,
            "name": meta.get("display_name") or disp.get(bb) or c.get("block_name") or bb,
            "score": score,
            "weight": c.get("weight_in_apex", c.get("weight_in_ppk")),
            "gate": bool(c.get("block_gate_triggered")),
            "notes": meta.get("aggregator_spec"),
        })
    return out


def _indicators_by_block(recs: Dict[str, Any]) -> Dict[str, List[dict]]:
    """Group indicators by block from either 07.indicators[] (flat: drone/base)
    or 07.points[].indicators[] (per-point: gcp/checkpoint, worst score wins)."""
    flat = recs.get("indicators")
    if isinstance(flat, list) and flat:
        out: Dict[str, List[dict]] = {}
        for i in flat:
            out.setdefault(i.get("block"), []).append(i)
        return out
    agg: Dict[str, Dict[str, dict]] = {}
    for p in (recs.get("points") or []):
        for i in (p.get("indicators") or []):
            bb, iid = i.get("block"), i.get("indicator_id")
            cur = agg.setdefault(bb, {}).get(iid)
            if cur is None or (i.get("score") is not None
                               and i.get("score") < (cur.get("score") or 1e9)):
                agg.setdefault(bb, {})[iid] = i
    return {bb: list(d.values()) for bb, d in agg.items()}


def _normalize_blocks(score06: Dict[str, Any], recs: Dict[str, Any],
                      agg05: Dict[str, Any], disp: Dict[str, str]) -> List[Dict[str, Any]]:
    """Uniform [{bb_id, name, score, weight, gate, notes, indicators[]}] across all
    three pipeline shapes (block source and indicator source resolved independently)."""
    by_block = _indicators_by_block(recs)
    blocks = _block_meta(score06, agg05, disp)
    for blk in blocks:
        blk["indicators"] = by_block.get(blk["bb_id"], [])
    return blocks


def build_result(sub: Subsystem) -> Dict[str, Any]:
    """Read this subsystem's outputs and return the hardware-page result object."""
    out = sub.outputs_dir
    score06 = _envelope_data(_load(out / Path(sub.score_file).name))
    recs = _envelope_data(_load(out / "07_recommendations.json"))

    # spec file (for full band ladders) — path lives in the codebase paths.json
    spec = {}
    paths = _load(sub.codebase / "paths.json") or {}
    spec_rel = paths.get("spec_file")
    if spec_rel:
        spec = _load(sub.codebase / spec_rel) or {}
    thr_by_ind = _thresholds_by_indicator(spec)
    spec_inds = {i.get("indicator_id"): i for i in (spec.get("indicators") or [])}

    apex = score06.get(sub.score_key)
    if apex is None:
        apex = recs.get("apex_score")
    decision = recs.get("decision") or "unable_to_assess"

    agg05 = _envelope_data(_load(out / "05_building_blocks.json")).get("aggregated_blocks") or {}
    spec_bbs = [b for b in (spec.get("building_blocks") or []) if isinstance(b, dict)]
    block_disp = {b.get("block_id"): b.get("display_name") for b in spec_bbs}
    block_desc = {b.get("block_id"): (b.get("question") or b.get("meaning") or b.get("intent") or "")
                  for b in spec_bbs}
    norm_blocks = _normalize_blocks(score06, recs, agg05, block_disp)

    # ---- building blocks (pipeline's real blocks) --------------------------
    bbs: List[Dict[str, Any]] = []
    for blk in norm_blocks:
        bb_id = blk["bb_id"]
        indicators = []
        block_has_alert = False
        for i in blk["indicators"]:
            iid = i.get("indicator_id")
            sc = i.get("score")
            grades = _build_grades(thr_by_ind.get(iid, []), sc)
            mb = i.get("matched_band") or {}
            if not grades and mb:
                # specs without a flat thresholds table (gcp/checkpoint): show the
                # single attained band from the recommendations engine.
                grades = [{
                    "l": (mb.get("level") or "current").title(),
                    "r": mb.get("label") or "",
                    "s": sc,
                    "current": True,
                    "flag": mb.get("flag"),
                }]
            alert = _indicator_alert(i)
            if alert:
                block_has_alert = True
            sind = spec_inds.get(iid, {})
            acts = i.get("actions")
            acts = acts if isinstance(acts, list) else ([acts] if acts else [])
            indicators.append({
                "id": iid,
                "name": sind.get("display_name") or i.get("name") or iid,
                "desc": sind.get("meaning") or i.get("verified_statement") or "",
                "sources": [],
                "currentScore": sc,
                "level": mb.get("level"),
                "bandLabel": mb.get("label"),
                "grades": grades,
                "rec": _indicator_rec(i),
                "impact": i.get("impact"),                       # native band impact text
                "actions": [str(a) for a in acts],               # native band action steps (→ list)
                "verifiedStatement": i.get("verified_statement"),
                "alert": alert,
            })
        bbs.append({
            "id": _short_id(bb_id),
            "bbId": bb_id,
            "name": blk["name"],
            "desc": block_desc.get(bb_id) or blk["notes"] or "",
            "score": round(blk["score"] or 0),
            "rawScore": blk["score"],
            "weight": blk["weight"],
            "anomaly": blk["gate"] or block_has_alert,
            "notes": blk["notes"],
            "indicators": indicators,
        })

    # ---- flat per-indicator scores for the page's confidence engine ---------
    # The hardware pages render from a scores{indicator_id: 0-100} object. 04 carries
    # the complete per-indicator list (single-record shape); fall back to the blocks.
    scores: Dict[str, Any] = {}
    ind04 = _envelope_data(_load(out / "04_indicators.json")).get("indicators")
    if isinstance(ind04, list):
        for i in ind04:
            iid, sv = i.get("indicator_id"), i.get("score")
            if iid is not None and sv is not None:
                scores[iid] = sv
    for blk in norm_blocks:                       # supplement (per-point aggregate)
        for i in blk["indicators"]:
            iid, sv = i.get("indicator_id"), i.get("score")
            if iid and sv is not None and iid not in scores:
                scores[iid] = sv

    # ---- per-point scores (gcp / checkpoint) for the point selector ----------
    points = []
    for p in (recs.get("points") or []):
        psc = {}
        for i in (p.get("indicators") or []):
            iid, sv = i.get("indicator_id"), i.get("score")
            if iid is not None and sv is not None:
                psc[iid] = sv
        points.append({
            "id": p.get("point_id") or f"P{len(points) + 1}",
            "device_type": p.get("device_type"),
            "scores": psc,
        })

    tier = recs.get("tier_interpretation") or ""
    grade = tier.split("(")[0].strip() if tier else None

    # real per-run reason (07.decision_rationale), with raw indicator ids made readable
    name_by_id = {iid: (spec_inds.get(iid, {}).get("display_name") or iid) for iid in scores}
    for i in (spec.get("indicators") or []):
        if i.get("indicator_id"):
            name_by_id.setdefault(i["indicator_id"], i.get("display_name") or i["indicator_id"])
    rec_text = recs.get("decision_rationale") or ""
    rec_text = re.sub(r"L3I_[A-Z]+_\d+",
                      lambda m: name_by_id.get(m.group(0), m.group(0)), rec_text)

    result = {
        "subsystem": sub.key,
        "module": sub.module,
        "label": sub.label,
        "overallScore": round(apex) if isinstance(apex, (int, float)) else None,
        "rawScore": apex,
        "grade": grade,
        "tier": tier,
        "decision": decision,
        "decisionLabel": DECISION_LABEL.get(decision, decision.replace("_", " ").upper()),
        "verdict": DECISION_VERDICT.get(decision, "review"),
        "recommendation": rec_text,
        "globalGate": (bool(score06.get("global_gate_triggered"))
                       or (apex == 0)
                       or bool((recs.get("indicator_rollup") or {}).get("hard_gate_points"))),
        "flags": recs.get("all_flags_aggregated") or score06.get("all_flags_aggregated") or [],
        "flagsBySeverity": recs.get("flags_by_severity") or {},
        "bbs": bbs,
        "scores": scores,        # {indicator_id: 0-100} for the page's confidence engine
        "points": points,        # per-point scores (gcp / checkpoint); [] otherwise
        "nulls": [],             # indicators with unavailable inputs (none surfaced yet)
        "xlsxUrl": f"/api/download/{sub.key}/xlsx",
        "provenanceUrl": f"/api/download/{sub.key}/provenance",
    }

    # ---- per-point summary (GCP / CheckPoint) ------------------------------
    if sub.per_point:
        points = recs.get("points") or []
        result["perPoint"] = [
            {
                "pointId": p.get("point_id") or f"P{idx + 1}",
                "decision": p.get("point_decision"),
                "rationale": p.get("point_rationale"),
                "role": p.get("device_role"),
                "deviceType": p.get("device_type"),
            }
            for idx, p in enumerate(points)
        ]
        summary = recs.get("subsystem_summary") or {}
        result["pointCount"] = (summary.get("effective_check_point_count")
                                or summary.get("n_points") or len(points) or None)

    return result
