#!/usr/bin/env python3
"""Generate a single-file HTML reference for the Check Point (RTK) provenance bundle.

Reads `check_point_confidence_score/check_point_confidence_score.json` (canonical) and writes
`check_point_confidence_score/check_point_provenance.html` — a browser-navigable mirror of the
xlsx with sticky nav, per-table search filter, deep-link ID anchors, and severity + stage
badges. Self-contained: inline CSS + tiny inline JS; no network dependencies.

Ported from the GCP generator with CheckPoint divergences (D3):
  1. Severity badge slug sanitized from the FIRST WORD (CheckPoint severities include compound
     labels with spaces/parens like "LOW (workflow)" / "MEDIUM (advisory)") so classes stay valid.
  2. Indicators carry an extra `data_availability_by_device_type` column.
  3. Apex split across 06 (check_point_score) + 06b (score meta); apex block-weight key is
     `weight_in_check_point_score`; weights list is `check_point_score_blocks`.
  4. _meta uses `counts` (not audit_counts) and has no changelog -> render design_summary + previous.
  5. Brand "Check Point Provenance (RTK)".
  6. Leaves the ad-hoc check_point_provenance_v1.html untouched; writes the canonical
     check_point_provenance.html.

Run:  /opt/anaconda3/bin/python3 scripts/build_provenance_html.py
"""
import html
import os as __os
import json
import sys
from pathlib import Path

ROOT = (Path(__os.environ["CBMI_PIPELINE_ROOT"]) if __os.environ.get("CBMI_PIPELINE_ROOT") else Path(__file__).resolve().parent.parent)
SPEC_PATH = ROOT / "check_point_confidence_score" / "check_point_confidence_score.json"
OUT_PATH = ROOT / "check_point_confidence_score" / "check_point_provenance.html"


# --------------------------------------------------------------------------- helpers
def esc(v):
    if v is None:
        return '<span class="muted">—</span>'
    if isinstance(v, bool):
        return '<span class="bool-yes">✓ true</span>' if v else '<span class="bool-no">✗ false</span>'
    if isinstance(v, (dict, list)):
        return f'<code>{html.escape(json.dumps(v, sort_keys=True))}</code>'
    s = html.escape(str(v))
    return s if s.strip() else '<span class="muted">—</span>'


def code(v):
    if v is None or v == "":
        return '<span class="muted">—</span>'
    return f'<code>{html.escape(str(v))}</code>'


def severity_badge(sev):
    """CheckPoint severities can be compound ('LOW (workflow)'); slug from the first word
    so the CSS class is always valid (sev-low / sev-medium / sev-advisory / ...)."""
    if not sev:
        return ""
    slug = sev.split()[0].lower()
    return f'<span class="badge sev-{html.escape(slug)}">{html.escape(sev)}</span>'


def stage_badge(stage):
    if not stage:
        return ""
    return f'<span class="badge stg-{html.escape(stage.replace("_", "-"))}">{html.escape(stage)}</span>'


def id_anchor(ident):
    return f'<a class="rowid" id="{html.escape(ident)}" href="#{html.escape(ident)}">{html.escape(ident)}</a>'


def ref(ident):
    return f'<a class="ref" href="#{html.escape(str(ident))}">{html.escape(str(ident))}</a>'


def table_head(cols):
    return "<thead><tr>" + "".join(f"<th>{html.escape(c)}</th>" for c in cols) + "</tr></thead>"


def table_body(rows):
    return "<tbody>" + "".join(rows) + "</tbody>"


def row(cells):
    return "<tr>" + "".join(f"<td>{c}</td>" for c in cells) + "</tr>"


def section(slug, title, subtitle, body):
    return f"""
<section id="{slug}">
  <header class="section-head">
    <h2>{html.escape(title)}</h2>
    <div class="section-subtitle">{html.escape(subtitle)}</div>
    <input class="filter" data-target="tbl-{slug}" placeholder="Filter rows in this sheet (type any text)…" />
  </header>
  <div class="table-wrap">
    <table id="tbl-{slug}" class="data-table">
      {body}
    </table>
  </div>
</section>
"""


def section_kv(slug, title, subtitle, inner):
    return f"""
<section id="{slug}">
  <header class="section-head">
    <h2>{html.escape(title)}</h2>
    <div class="section-subtitle">{html.escape(subtitle)}</div>
  </header>
  {inner}
</section>
"""


def kv_table(pairs, slug=None):
    rows = [row([f'<strong>{html.escape(str(k))}</strong>', esc(v)]) for k, v in pairs]
    idattr = f' id="tbl-{slug}"' if slug else ""
    return f'<table class="data-table"{idattr}>{table_head(["field", "value"])}{table_body(rows)}</table>'


# --------------------------------------------------------------------------- sheet renderers
def render_readme(meta):
    it = ['<div class="hero-meta">']
    for label, key in [("Version", "version"), ("Workflow", "workflow"), ("Phase", "phase"),
                       ("Generated", "generated_at"), ("Previous", "previous")]:
        it.append(f'<div class="meta-pill">{label} <strong>{html.escape(str(meta.get(key, "")))}</strong></div>')
    it.append('</div>')
    it.append(f'<p class="desc">{html.escape(meta.get("design_summary", ""))}</p>')
    counts = meta.get("counts", {})
    it.append('<h3>Audit Counts</h3><div class="audit-grid">')
    for k, v in counts.items():
        it.append(f'<div class="audit-card"><div class="audit-n">{v}</div><div class="audit-k">{html.escape(k)}</div></div>')
    it.append('</div>')
    if meta.get("runtime_independence"):
        it.append('<h3>Runtime Independence</h3>')
        it.append(f'<p>{esc(meta["runtime_independence"])}</p>')
    pref = meta.get("id_prefixes", {})
    if pref:
        it.append('<h3>ID Prefixes</h3><ul class="bullet">')
        for k, v in pref.items():
            it.append(f'<li><code>{html.escape(k)}</code> — {html.escape(str(v))}</li>')
        it.append('</ul>')
    return f"""
<section id="readme">
  <header class="section-head">
    <h2>00 Readme &amp; Meta</h2>
    <div class="section-subtitle">{html.escape(meta.get("title", ""))}</div>
  </header>
  {''.join(it)}
</section>
"""


def render_source_files(spec):
    rows = [row([
        id_anchor(s["file_id"]), esc(s["file_name"]), esc(s.get("subsystem")),
        code(s.get("file_extensions")), esc(s.get("physical_source")), esc(s.get("acquisition_mode")),
        esc(s.get("is_aspirational")), esc(s.get("how_to_obtain")), esc(s.get("notes")),
    ]) for s in spec["source_files"]]
    body = table_head(["file_id", "file_name", "subsystem", "extensions", "physical_source",
                       "acquisition_mode", "is_aspirational", "how_to_obtain", "notes"]) + table_body(rows)
    return section("source-files", f"01 Source Files ({len(spec['source_files'])})",
                   "Physical sources from which per-point fields are extracted (RTK export / oplog / form).", body)


def render_source_fields(spec):
    rows = [row([
        id_anchor(f["field_id"]), esc(f["field_name"]), ref(f["file_id"]), esc(f.get("category")),
        esc(f.get("acquisition_mode")), esc(f.get("is_aspirational")), esc(f.get("meaning")), esc(f.get("notes")),
    ]) for f in spec["source_fields"]]
    body = table_head(["field_id", "field_name", "file_id", "category", "acquisition_mode",
                       "is_aspirational", "meaning", "notes"]) + table_body(rows)
    return section("source-fields", f"02 Source Fields ({len(spec['source_fields'])})",
                   "L1F_CP_* — extracted by Stage 2 parsers (per point).", body)


def render_derived_fields(spec):
    rows = [row([
        id_anchor(d["derived_id"]), esc(d["derived_name"]), esc(d.get("kind")),
        code(d.get("formula_expression")), code(d.get("input_field_names")), esc(d.get("meaning")),
    ]) for d in spec["derived_fields"]]
    body = table_head(["derived_id", "derived_name", "kind", "formula_expression",
                       "input_field_names", "meaning"]) + table_body(rows)
    return section("derived-fields", f"03 Derived Fields ({len(spec['derived_fields'])})",
                   "L2D_CP_* — computed per point in Stage 3a (15 per-point + 1 survey-level).", body)


def render_indicators(spec):
    rows = []
    for i in spec["indicators"]:
        gate = ('<span class="badge gate">internal gate</span> '
                if str(i.get("has_internal_gate")).upper() in ("TRUE", "YES") or i.get("has_internal_gate") is True
                else "")
        rows.append(row([
            id_anchor(i["indicator_id"]), esc(i["indicator_name"]), esc(i.get("display_name")),
            ref(i["building_block_id"]), esc(i.get("weight_in_block")), code(i.get("input_derived_fields")),
            esc(i.get("covers_problems")), gate + esc(i.get("gate_condition") or ""),
            esc(i.get("gate_action") or ""), esc(i.get("threshold_summary")),
            esc(i.get("data_availability_by_device_type")), esc(i.get("justification")),
        ]))
    body = table_head(["indicator_id", "indicator_name", "display_name", "building_block_id",
                       "weight_in_block", "input_derived_fields", "covers_problems", "internal_gate",
                       "gate_action", "threshold_summary (bands)", "data_availability_by_device_type",
                       "justification"]) + table_body(rows)
    return section("indicators", f"04 Indicators ({len(spec['indicators'])})",
                   "L3I_CP_* — scored in Stage 3b; band ladder is in threshold_summary (Option B). "
                   "data_availability_by_device_type captures RTK device conditionality.", body)


def render_building_blocks(spec):
    rows = []
    for b in spec["building_blocks"]:
        gate = ('<span class="badge gate">internal gate</span> '
                if b.get("has_internal_gate") in (True, "TRUE", "YES") else "")
        rows.append(row([
            id_anchor(b["block_id"]), esc(b["block_name"]), esc(b.get("display_name")),
            f'<strong>{esc(b.get("weight_in_check_point_score"))}</strong>', esc(b.get("question")),
            esc(b.get("failure_owner")), esc(b.get("operator_action")), esc(b.get("aggregator")),
            code(b.get("indicators_within")), gate + esc(b.get("gate_condition") or ""),
            esc(b.get("gate_action") or ""),
        ]))
    body = table_head(["block_id", "block_name", "display_name", "weight_in_check_point_score", "question",
                       "failure_owner", "operator_action", "aggregator", "indicators_within",
                       "internal_gate", "gate_action"]) + table_body(rows)
    return section("building-blocks", f"05 Building Blocks ({len(spec['building_blocks'])})",
                   "BB_CP_* — rolled up per point then cross-point aggregated in Stage 3c.", body)


def render_score(spec):
    g = spec["check_point_score"]
    kv = [(k, g.get(k)) for k in ["score_id", "display_name", "workflow", "phase",
          "formula_expression", "per_point_formula", "aggregator", "global_gate_condition",
          "global_gate_action", "null_handling"]]
    wrows = [row([ref(w["block_id"]), esc(w.get("block_name")), f'<strong>{esc(w.get("weight"))}</strong>'])
             for w in spec.get("check_point_score_blocks", [])]
    weights = f'<table class="data-table">{table_head(["block_id", "block_name", "weight"])}{table_body(wrows)}</table>'
    inner = (f'<h3 class="sub-h3">Apex scoring</h3>{kv_table(kv, slug="check-point-score")}'
             f'<h3 class="sub-h3">Apex block weights (sum = 1.0)</h3>{weights}')
    return section_kv("check-point-score", "06 Check Point Score",
                      "Apex formula, per-point formula, gates, and block weights.", inner)


def render_score_meta(spec):
    g = spec["check_point_score"]
    kv = [(k, g.get(k)) for k in ["scope_note", "source_file_set", "device_types_supported", "known_limitations"]]
    return section_kv("score-meta", "06b Score Meta", "Apex scope, supported devices, and known limitations.",
                      kv_table(kv, slug="score-meta"))


def render_flags(spec):
    rows = [row([
        id_anchor(f["flag_id"]), esc(f["flag_name"]), esc(f.get("raised_by_type")),
        stage_badge(f.get("raised_at_stage")), esc(f.get("condition")), severity_badge(f.get("severity")),
        esc(f.get("covers_problems")), esc(f.get("meaning")),
    ]) for f in spec["flags"]]
    body = table_head(["flag_id", "flag_name", "raised_by_type", "raised_at_stage", "condition",
                       "severity", "covers_problems", "meaning"]) + table_body(rows)
    return section("flags", f"07 Flags ({len(spec['flags'])})",
                   "FLG_CP_* — grouped by raised_at_stage. Severity + stage badges.", body)


def render_problem_coverage(spec):
    rows = []
    for p in spec["problem_coverage_map"]:
        pid = f"problem-{p['problem_no']}"
        rows.append(row([
            f'<a class="rowid" id="{pid}" href="#{pid}">#{p["problem_no"]}</a>',
            esc(p.get("problem")), severity_badge(p.get("severity")), esc(p.get("frequency")),
            code(p.get("disposition")), esc(p.get("covered_by")),
        ]))
    body = table_head(["problem_no", "problem", "severity", "frequency", "disposition",
                       "covered_by"]) + table_body(rows)
    return section("problem-coverage", f"08 Problem Coverage ({len(spec['problem_coverage_map'])})",
                   "CBMI Check Point problems and how this subsystem (or a sibling) covers each.", body)


# --------------------------------------------------------------------------- nav / css / js
NAV_LINKS = [
    ("readme", "Readme"), ("source-files", "01 Source Files"), ("source-fields", "02 Source Fields"),
    ("derived-fields", "03 Derived Fields"), ("indicators", "04 Indicators"),
    ("building-blocks", "05 Building Blocks"), ("check-point-score", "06 Check Point Score"),
    ("score-meta", "06b Score Meta"), ("flags", "07 Flags"), ("problem-coverage", "08 Problem Coverage"),
]


def nav():
    items = "".join(f'<a href="#{slug}">{html.escape(label)}</a>' for slug, label in NAV_LINKS)
    return f'<nav class="topnav"><div class="brand">Check Point Provenance (RTK)</div><div class="nav-links">{items}</div></nav>'


CSS = """
:root{--bg:#0f172a;--panel:#fff;--ink:#0f172a;--ink-soft:#475569;--line:#e5e7eb;--accent:#1e40af;--accent-soft:#dbeafe;--code-bg:#f1f5f9;--table-stripe:#f8fafc;--table-hover:#eff6ff;}
*{box-sizing:border-box;}
body{margin:0;font:14px/1.45 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;color:var(--ink);background:#f1f5f9;}
.topnav{position:sticky;top:0;z-index:50;background:var(--bg);color:#fff;padding:10px 24px;display:flex;align-items:center;gap:24px;box-shadow:0 2px 8px rgba(0,0,0,.15);flex-wrap:wrap;}
.topnav .brand{font-size:16px;font-weight:700;letter-spacing:.3px;margin-right:12px;}
.topnav .nav-links{display:flex;gap:4px;flex-wrap:wrap;}
.topnav .nav-links a{color:#cbd5e1;text-decoration:none;padding:6px 10px;border-radius:6px;font-size:12.5px;}
.topnav .nav-links a:hover{background:rgba(255,255,255,.1);color:#fff;}
main{max-width:1500px;margin:0 auto;padding:24px;}
section{background:var(--panel);padding:20px 24px 28px;margin-bottom:24px;border-radius:10px;box-shadow:0 1px 4px rgba(15,23,42,.07);}
.section-head{margin-bottom:14px;}
.section-head h2{margin:0 0 4px 0;font-size:22px;color:var(--ink);}
.section-subtitle{color:var(--ink-soft);font-size:13px;margin-bottom:12px;}
.sub-h3{margin-top:18px;font-size:15px;color:var(--ink);}
input.filter{width:100%;padding:8px 12px;font:13px/1.4 -apple-system,sans-serif;border:1px solid var(--line);border-radius:6px;margin-top:6px;}
input.filter:focus{outline:none;border-color:var(--accent);box-shadow:0 0 0 3px var(--accent-soft);}
.table-wrap{overflow-x:auto;margin-top:4px;}
table.data-table{width:100%;border-collapse:collapse;font-size:12.5px;margin-bottom:6px;}
table.data-table th{background:#1e293b;color:#fff;padding:8px 10px;text-align:left;font-weight:600;font-size:11.5px;text-transform:uppercase;letter-spacing:.4px;position:sticky;top:56px;}
table.data-table td{padding:7px 10px;border-bottom:1px solid var(--line);vertical-align:top;font-size:12.5px;}
table.data-table tbody tr:nth-child(even){background:var(--table-stripe);}
table.data-table tbody tr:hover{background:var(--table-hover);}
a.rowid{font-family:ui-monospace,"SF Mono",Menlo,monospace;font-weight:600;color:var(--accent);text-decoration:none;border-bottom:1px dotted var(--accent);}
a.rowid:hover{background:var(--accent-soft);}
a.ref{color:var(--accent);text-decoration:none;font-family:ui-monospace,"SF Mono",Menlo,monospace;font-size:12px;}
a.ref:hover{text-decoration:underline;}
code{font:12px/1.4 ui-monospace,"SF Mono",Menlo,monospace;background:var(--code-bg);padding:1px 5px;border-radius:3px;color:#0f172a;white-space:pre-wrap;word-break:break-word;}
.muted{color:#94a3b8;}
.bool-yes{color:#16a34a;font-weight:600;}
.bool-no{color:#94a3b8;}
.badge{display:inline-block;padding:1px 8px;border-radius:999px;font-size:11px;font-weight:600;letter-spacing:.3px;background:#e2e8f0;color:#475569;}
.sev-catastrophic{background:#fca5a5;color:#7f1d1d;}
.sev-critical{background:#fee2e2;color:#b91c1c;}
.sev-high{background:#ffedd5;color:#c2410c;}
.sev-medium{background:#fef3c7;color:#b45309;}
.sev-low{background:#dbeafe;color:#1d4ed8;}
.sev-advisory{background:#e0e7ff;color:#4338ca;}
.stg-threshold{background:#fef9c3;color:#854d0e;}
.stg-internal-gate{background:#fed7aa;color:#9a3412;}
.stg-global-gate{background:#fecaca;color:#991b1b;}
.stg-composite{background:#cffafe;color:#155e75;}
.stg-null-handler{background:#e2e8f0;color:#475569;}
.badge.gate{background:#fde68a;color:#92400e;}
.hero-meta{display:flex;flex-wrap:wrap;gap:10px;margin:8px 0 16px 0;}
.meta-pill{background:var(--code-bg);padding:6px 12px;border-radius:999px;font-size:12.5px;}
.meta-pill strong{color:var(--accent);}
p.desc{color:var(--ink-soft);}
.audit-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:10px;margin:8px 0 18px 0;}
.audit-card{background:var(--code-bg);padding:10px 12px;border-radius:8px;text-align:center;}
.audit-n{font-size:22px;font-weight:700;color:var(--accent);}
.audit-k{font-size:11.5px;color:var(--ink-soft);text-transform:uppercase;letter-spacing:.5px;margin-top:2px;}
ul.bullet{padding-left:18px;margin:4px 0 12px;}
ul.bullet li{margin:4px 0;}
ul.bullet code{font-size:11.5px;}
@media (max-width:380px){table.data-table{font-size:11px;}}
"""

JS = """
document.addEventListener("DOMContentLoaded",()=>{
  document.querySelectorAll("input.filter").forEach((inp)=>{
    const tbl=document.getElementById(inp.dataset.target);
    if(!tbl)return;
    inp.addEventListener("input",()=>{
      const q=inp.value.trim().toLowerCase();
      tbl.querySelectorAll("tbody tr").forEach((tr)=>{
        tr.style.display=q&&!tr.textContent.toLowerCase().includes(q)?"none":"";
      });
    });
  });
  document.querySelectorAll('a[href^="#"]').forEach((a)=>{
    a.addEventListener("click",(ev)=>{
      const id=a.getAttribute("href").slice(1);
      const el=document.getElementById(id);
      if(!el)return;
      ev.preventDefault();
      const top=el.getBoundingClientRect().top+window.pageYOffset-80;
      window.scrollTo({top,behavior:"smooth"});
      if(history.pushState)history.pushState(null,"","#"+id);
    });
  });
});
"""


def main():
    spec = json.loads(SPEC_PATH.read_text())
    meta = spec["_meta"]
    sheets_html = "\n".join([
        render_readme(meta), render_source_files(spec), render_source_fields(spec),
        render_derived_fields(spec), render_indicators(spec), render_building_blocks(spec),
        render_score(spec), render_score_meta(spec), render_flags(spec), render_problem_coverage(spec),
    ])
    title = f"{meta.get('title', 'check_point_confidence_score')} v{meta.get('version', '')}"
    out = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{html.escape(title)} — Reference</title>
<style>{CSS}</style>
</head>
<body>
{nav()}
<main>
{sheets_html}
</main>
<footer style="text-align:center;padding:20px 0;color:#64748b;font-size:12px;">
  Generated from <code>check_point_confidence_score.json</code> v{html.escape(str(meta.get('version', '')))}
  &middot; This HTML mirrors <code>check_point_confidence_score.xlsx</code>; the JSON is the canonical source.
</footer>
<script>{JS}</script>
</body>
</html>
"""
    OUT_PATH.write_text(out)
    try:
        shown = OUT_PATH.relative_to(ROOT)
    except ValueError:
        shown = OUT_PATH
    print(f"wrote {shown}")
    print(f"   size: {OUT_PATH.stat().st_size:,} bytes")
    print(f"   sections: {len(NAV_LINKS)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
