#!/usr/bin/env python3
"""Generate a single-file HTML reference for the Drone PPK provenance bundle.

Reads `drone_provenance_ppk/drone_provenance_ppk.json` (canonical) and writes
`drone_provenance_ppk/drone_provenance_ppk.html` — a browser-navigable mirror
of the xlsx with:

- Sticky top nav linking to every sheet
- Per-table search filter (no external JS deps)
- ID anchors (#L3I_IMG_001 / #BB_IMG_CAPTURE / #FLG_017 etc.) for direct linking
- Color-coded severity badges and stage badges on the flags sheet
- All meta + readme + audit counts at the top
- Self-contained: inline CSS + tiny inline JS; no network dependencies

Run:  python3 scripts/build_provenance_html.py
"""
import html
import os as __os
import json
import sys
from pathlib import Path


SPEC_PATH = (Path(__os.environ["CBMI_PIPELINE_ROOT"]) if __os.environ.get("CBMI_PIPELINE_ROOT") else Path(__file__).resolve().parent.parent) / "drone_provenance_ppk" / "drone_provenance_ppk.json"
OUT_PATH  = (Path(__os.environ["CBMI_PIPELINE_ROOT"]) if __os.environ.get("CBMI_PIPELINE_ROOT") else Path(__file__).resolve().parent.parent) / "drone_provenance_ppk" / "drone_provenance_ppk.html"


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

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


def severity_badge(sev: str) -> str:
    if not sev:
        return ""
    cls = f"sev-{sev.lower()}"
    return f'<span class="badge {cls}">{html.escape(sev)}</span>'


def stage_badge(stage: str) -> str:
    if not stage:
        return ""
    cls = f"stg-{stage.replace('_', '-')}"
    return f'<span class="badge {cls}">{html.escape(stage)}</span>'


def id_anchor(prefix: str, ident: str) -> str:
    return f'<a class="rowid" id="{html.escape(ident)}" href="#{html.escape(ident)}">{html.escape(ident)}</a>'


def section(slug: str, title: str, subtitle: str, body: str) -> str:
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


def table_head(cols: list[str]) -> str:
    return "<thead><tr>" + "".join(f"<th>{html.escape(c)}</th>" for c in cols) + "</tr></thead>"


def table_body(rows: list[str]) -> str:
    return "<tbody>" + "".join(rows) + "</tbody>"


def row(cells: list[str]) -> str:
    return "<tr>" + "".join(f"<td>{c}</td>" for c in cells) + "</tr>"


# ---------------------------------------------------------------------------
# sheet renderers
# ---------------------------------------------------------------------------

def render_readme(meta: dict, readme: dict) -> str:
    items = []
    # Key headline
    items.append('<div class="hero-meta">')
    items.append(f'<div class="meta-pill">Version <strong>{html.escape(str(meta.get("version", "")))}</strong></div>')
    items.append(f'<div class="meta-pill">Workflow <strong>{html.escape(str(meta.get("workflow", "")))}</strong></div>')
    items.append(f'<div class="meta-pill">Generated <strong>{html.escape(str(meta.get("generated_at", "")))}</strong></div>')
    items.append(f'<div class="meta-pill">Previous <strong>{html.escape(str(meta.get("previous_version", "")))}</strong></div>')
    items.append('</div>')

    # description
    items.append(f'<p class="desc">{html.escape(meta.get("description", ""))}</p>')

    # audit counts
    counts = meta.get("audit_counts", {})
    items.append('<h3>Audit Counts</h3>')
    items.append('<div class="audit-grid">')
    for k, v in counts.items():
        items.append(f'<div class="audit-card"><div class="audit-n">{v}</div><div class="audit-k">{html.escape(k)}</div></div>')
    items.append('</div>')

    # critical gates
    cg = meta.get("critical_gates", {})
    if cg:
        items.append('<h3>Critical Gates</h3><ul class="bullet">')
        for k, v in cg.items():
            items.append(f'<li><code>{html.escape(k)}</code>: {html.escape(v)}</li>')
        items.append('</ul>')

    # notes
    for k in ("cal_conf_note", "bin_consolidation_note", "user_input_form_note", "external_api_note"):
        if k in meta:
            items.append(f'<h3>{html.escape(k.replace("_", " ").title())}</h3>')
            items.append(f'<p>{html.escape(meta[k])}</p>')

    # changelogs
    for k in sorted(k for k in meta if k.startswith("changelog_")):
        items.append(f'<h3>{html.escape(k.replace("_", " "))}</h3><ul class="bullet">')
        for ln in meta[k]:
            items.append(f'<li>{html.escape(ln)}</li>')
        items.append('</ul>')

    # readme.RECOMMENDED READ ORDER FOR AN AGENT
    rr = readme.get("RECOMMENDED READ ORDER FOR AN AGENT")
    if rr:
        items.append('<h3>Recommended Read Order</h3>')
        items.append(f'<p>{html.escape(rr)}</p>')

    return f"""
<section id="readme">
  <header class="section-head">
    <h2>00 Readme &amp; Meta</h2>
    <div class="section-subtitle">{html.escape(meta.get("title", ""))}</div>
  </header>
  {''.join(items)}
</section>
"""


def render_source_files(spec) -> str:
    rows = []
    for s in spec["source_files"]:
        rows.append(row([
            id_anchor("src", s["file_id"]),
            esc(s["file_name"]),
            esc(s["subsystem"]),
            code(s["file_extensions"]),
            esc(s["physical_source"]),
            esc(s["is_aspirational"]),
            esc(s.get("how_to_obtain")),
            esc(s.get("notes")),
        ]))
    body = table_head(["file_id", "file_name", "subsystem", "extensions", "physical_source", "is_aspirational", "how_to_obtain", "notes"]) + table_body(rows)
    return section("source-files", "01 Source Files (6)",
                   "Physical sources from which fields are extracted.", body)


def render_source_fields(spec) -> str:
    rows = []
    for f in spec["source_fields"]:
        rows.append(row([
            id_anchor("fld", f["field_id"]),
            esc(f["field_name"]),
            f'<a class="ref" href="#{f["file_id"]}">{html.escape(f["file_id"])}</a>',
            esc(f["category"]),
            code(f["data_type"]),
            esc(f.get("units")),
            esc(f["acquisition_mode"]),
            esc(f["is_aspirational"]),
            esc(f.get("meaning")),
            esc(f.get("notes")),
        ]))
    body = table_head(["field_id", "field_name", "file_id", "category", "data_type", "units",
                       "acquisition_mode", "is_aspirational", "meaning", "notes"]) + table_body(rows)
    return section("source-fields", "02 Source Fields (88)",
                   "L1F_* fields extracted by Stage 2 parsers.", body)


def render_derived_fields(spec) -> str:
    rows = []
    for d in spec["derived_fields"]:
        rows.append(row([
            id_anchor("der", d["derived_id"]),
            esc(d["derived_name"]),
            esc(d["subsystem"]),
            esc(d.get("units")),
            code(d.get("formula_expression")),
            esc(d.get("formula_notes")),
            code(d.get("input_field_ids")),
            esc(d.get("meaning")),
        ]))
    body = table_head(["derived_id", "derived_name", "subsystem", "units",
                       "formula_expression", "formula_notes", "input_field_ids", "meaning"]) + table_body(rows)
    return section("derived-fields", "03 Derived Fields (32)",
                   "L2D_* — computed in Stage 3a from L1F_* sources and/or other L2D_*.", body)


def render_indicators(spec) -> str:
    rows = []
    for i in spec["indicators"]:
        gate = ""
        if i.get("has_internal_gate"):
            gate = '<span class="badge gate">internal gate</span>'
        rows.append(row([
            id_anchor("ind", i["indicator_id"]),
            esc(i["indicator_name"]),
            esc(i["display_name"]),
            f'<a class="ref" href="#{i["building_block_id"]}">{html.escape(i["building_block_id"])}</a>',
            esc(i["weight_in_block"]),
            code(i.get("input_derived_field_ids")),
            code(i.get("input_source_field_ids")),
            esc(i.get("meaning")),
            gate + " " + esc(i.get("gate_condition") or ""),
            esc(i.get("gate_action") or ""),
        ]))
    body = table_head(["indicator_id", "indicator_name", "display_name", "building_block_id",
                       "weight_in_block", "input_derived_field_ids", "input_source_field_ids",
                       "meaning", "internal_gate", "gate_action"]) + table_body(rows)
    return section("indicators", "04 Indicators (23)",
                   "L3I_* — scored in Stage 3b against threshold bands.", body)


def render_thresholds(spec) -> str:
    rows = []
    for t in spec["thresholds"]:
        flag = t.get("flag_raised")
        flag_cell = f'<a class="ref" href="#flg-anchor-{html.escape(flag)}">{html.escape(flag)}</a>' if flag else esc(None)
        rows.append(row([
            id_anchor("th", t["threshold_id"]),
            f'<a class="ref" href="#{t["indicator_id"]}">{html.escape(t["indicator_id"])}</a>',
            esc(t["band_order"]),
            code(t.get("condition_expression")),
            esc(t.get("condition_text")),
            f'<span class="score-pill">{t["score_value"]}</span>',
            flag_cell,
        ]))
    body = table_head(["threshold_id", "indicator_id", "band_order",
                       "condition_expression", "condition_text", "score_value", "flag_raised"]) + table_body(rows)
    return section("thresholds", "05 Thresholds (100)",
                   "Per-indicator band ladder. First-match-wins, evaluated top-down.", body)


def render_building_blocks(spec) -> str:
    rows = []
    for b in spec["building_blocks"]:
        gate = ""
        if b.get("has_internal_gate"):
            gate = '<span class="badge gate">internal gate</span>'
        rows.append(row([
            id_anchor("bb", b["block_id"]),
            esc(b["block_name"]),
            esc(b["display_name"]),
            esc(b["purpose"]),
            f'<strong>{b["weight_in_drone_score_ppk"]}</strong>',
            code(b.get("formula_expression")),
            gate + " " + esc(b.get("gate_condition") or ""),
            esc(b.get("gate_action") or ""),
            esc(b.get("disabled_in_workflows") or ""),
            esc(b.get("weight_redistribution_rule") or ""),
        ]))
    body = table_head(["block_id", "block_name", "display_name", "purpose",
                       "weight_in_drone_score_ppk", "formula_expression",
                       "internal_gate", "gate_action", "disabled_in_workflows",
                       "weight_redistribution_rule"]) + table_body(rows)
    return section("building-blocks", "06 Building Blocks (4)",
                   "Rolled up in Stage 3c. CAL_CONF has weight 0 — parallel deliverable.", body)


def render_block_composition(spec) -> str:
    rows = []
    for c in spec["block_composition"]:
        rows.append(row([
            f'<a class="ref" href="#{c["block_id"]}">{html.escape(c["block_id"])}</a>',
            f'<a class="ref" href="#{c["indicator_id"]}">{html.escape(c["indicator_id"])}</a>',
            f'<strong>{c["weight"]}</strong>',
        ]))
    body = table_head(["block_id", "indicator_id", "weight"]) + table_body(rows)
    return section("block-composition", "07 Block Composition (23)",
                   "Many-to-many block ↔ indicator weight table — authoritative for Stage 3c sums.", body)


def render_drone_score(spec) -> str:
    ds = spec["drone_score"]
    meta = ds["metadata"]
    rows_meta = []
    for k, v in meta.items():
        rows_meta.append(row([
            f'<strong>{html.escape(k)}</strong>',
            esc(v),
        ]))
    meta_block = table_head(["field", "value"]) + table_body(rows_meta)
    weights_rows = []
    for w in ds["weights"]:
        weights_rows.append(row([
            f'<a class="ref" href="#{w["block_id"]}">{html.escape(w["block_id"])}</a>',
            f'<strong>{w["weight_in_ppk"]}</strong>',
            esc(w.get("notes")),
        ]))
    weights_block = table_head(["block_id", "weight_in_ppk", "notes"]) + table_body(weights_rows)
    body = f"""
<h3 class="sub-h3">Metadata</h3>
{meta_block}
<h3 class="sub-h3">PPK Weights</h3>
{weights_block}
"""
    return section("drone-score", "08 Drone Score",
                   "Apex formula and per-block PPK weight; sums to 1.0.",
                   body)


def render_flags(spec) -> str:
    rows = []
    for f in spec["flags"]:
        # Add a target anchor for thresholds to link in
        rows.append(row([
            id_anchor("flg", f["flag_id"]) + f'<a class="hidden" id="flg-anchor-{html.escape(f["flag_name"])}"></a>',
            esc(f["flag_name"]),
            esc(f.get("raised_by_id")),
            esc(f.get("raised_by_type")),
            stage_badge(f.get("raised_at_stage")),
            esc(f.get("condition")),
            severity_badge(f.get("severity")),
            esc(f.get("meaning")),
        ]))
    body = table_head(["flag_id", "flag_name", "raised_by_id", "raised_by_type",
                       "raised_at_stage", "condition", "severity", "meaning"]) + table_body(rows)
    return section("flags", "09 Flags (19)",
                   "Grouped by raised_at_stage. CRITICAL/HIGH/MEDIUM/LOW severity badges.", body)


def render_script_hints(spec) -> str:
    rows = []
    for h in spec.get("script_hints", []):
        rows.append(row([
            f'<a class="ref" href="#{h["indicator_id"]}">{html.escape(h["indicator_id"])}</a>',
            esc(h.get("indicator_name")),
            f'<pre class="pseudocode">{html.escape(h.get("pseudocode") or "")}</pre>',
        ]))
    body = table_head(["indicator_id", "indicator_name", "pseudocode"]) + table_body(rows)
    return section("script-hints", "10 Script Hints (24)",
                   "Python-style templates implementing each indicator's threshold ladder.", body)


# ---------------------------------------------------------------------------
# top nav + style + script
# ---------------------------------------------------------------------------

NAV_LINKS = [
    ("readme", "Readme"),
    ("source-files", "01 Source Files"),
    ("source-fields", "02 Source Fields"),
    ("derived-fields", "03 Derived Fields"),
    ("indicators", "04 Indicators"),
    ("thresholds", "05 Thresholds"),
    ("building-blocks", "06 Building Blocks"),
    ("block-composition", "07 Block Composition"),
    ("drone-score", "08 Drone Score"),
    ("flags", "09 Flags"),
    ("script-hints", "10 Script Hints"),
]


def nav() -> str:
    items = "".join(f'<a href="#{slug}">{html.escape(label)}</a>' for slug, label in NAV_LINKS)
    return f'<nav class="topnav"><div class="brand">Drone Provenance PPK</div><div class="nav-links">{items}</div></nav>'


CSS = """
:root {
  --bg: #0f172a;
  --panel: #ffffff;
  --ink: #0f172a;
  --ink-soft: #475569;
  --line: #e5e7eb;
  --accent: #1e40af;
  --accent-soft: #dbeafe;
  --code-bg: #f1f5f9;
  --table-stripe: #f8fafc;
  --table-hover: #eff6ff;
}

* { box-sizing: border-box; }

body {
  margin: 0;
  font: 14px/1.45 -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
  color: var(--ink);
  background: #f1f5f9;
}

.topnav {
  position: sticky; top: 0; z-index: 50;
  background: var(--bg);
  color: white;
  padding: 10px 24px;
  display: flex; align-items: center; gap: 24px;
  box-shadow: 0 2px 8px rgba(0,0,0,0.15);
  flex-wrap: wrap;
}
.topnav .brand {
  font-size: 16px; font-weight: 700; letter-spacing: 0.3px;
  margin-right: 12px;
}
.topnav .nav-links {
  display: flex; gap: 4px; flex-wrap: wrap;
}
.topnav .nav-links a {
  color: #cbd5e1; text-decoration: none;
  padding: 6px 10px; border-radius: 6px;
  font-size: 12.5px;
}
.topnav .nav-links a:hover {
  background: rgba(255,255,255,0.1); color: white;
}

main {
  max-width: 1500px;
  margin: 0 auto;
  padding: 24px;
}

section {
  background: var(--panel);
  padding: 20px 24px 28px;
  margin-bottom: 24px;
  border-radius: 10px;
  box-shadow: 0 1px 4px rgba(15, 23, 42, 0.07);
}

.section-head { margin-bottom: 14px; }
.section-head h2 {
  margin: 0 0 4px 0;
  font-size: 22px;
  color: var(--ink);
}
.section-subtitle { color: var(--ink-soft); font-size: 13px; margin-bottom: 12px; }
.sub-h3 { margin-top: 18px; font-size: 15px; color: var(--ink); }

input.filter {
  width: 100%;
  padding: 8px 12px;
  font: 13px/1.4 -apple-system, sans-serif;
  border: 1px solid var(--line);
  border-radius: 6px;
  margin-top: 6px;
}
input.filter:focus {
  outline: none;
  border-color: var(--accent);
  box-shadow: 0 0 0 3px var(--accent-soft);
}

.table-wrap { overflow-x: auto; margin-top: 4px; }
table.data-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 12.5px;
}
table.data-table th {
  background: #1e293b;
  color: white;
  padding: 8px 10px;
  text-align: left;
  font-weight: 600;
  font-size: 11.5px;
  text-transform: uppercase;
  letter-spacing: 0.4px;
  position: sticky; top: 56px;
}
table.data-table td {
  padding: 7px 10px;
  border-bottom: 1px solid var(--line);
  vertical-align: top;
  font-size: 12.5px;
}
table.data-table tbody tr:nth-child(even) { background: var(--table-stripe); }
table.data-table tbody tr:hover { background: var(--table-hover); }

a.rowid {
  font-family: ui-monospace, "SF Mono", Menlo, monospace;
  font-weight: 600;
  color: var(--accent);
  text-decoration: none;
  border-bottom: 1px dotted var(--accent);
}
a.rowid:hover { background: var(--accent-soft); }

a.ref {
  color: var(--accent);
  text-decoration: none;
  font-family: ui-monospace, "SF Mono", Menlo, monospace;
  font-size: 12px;
}
a.ref:hover { text-decoration: underline; }

code {
  font: 12px/1.4 ui-monospace, "SF Mono", Menlo, monospace;
  background: var(--code-bg);
  padding: 1px 5px;
  border-radius: 3px;
  color: #0f172a;
  white-space: pre-wrap;
  word-break: break-word;
}

pre.pseudocode {
  margin: 0;
  font: 11.5px/1.4 ui-monospace, "SF Mono", Menlo, monospace;
  background: var(--code-bg);
  padding: 8px 10px;
  border-radius: 4px;
  max-height: 220px;
  overflow: auto;
  white-space: pre;
}

.muted { color: #94a3b8; }
.bool-yes { color: #16a34a; font-weight: 600; }
.bool-no { color: #94a3b8; }

.score-pill {
  display: inline-block;
  background: var(--accent-soft);
  color: var(--accent);
  padding: 1px 9px;
  border-radius: 999px;
  font-weight: 600;
  font-size: 12px;
}

.badge {
  display: inline-block;
  padding: 1px 8px;
  border-radius: 999px;
  font-size: 11px;
  font-weight: 600;
  letter-spacing: 0.3px;
}
.sev-critical { background: #fee2e2; color: #b91c1c; }
.sev-high     { background: #ffedd5; color: #c2410c; }
.sev-medium   { background: #fef3c7; color: #b45309; }
.sev-low      { background: #dbeafe; color: #1d4ed8; }

.stg-pre-score-ingestion { background: #e2e8f0; color: #475569; }
.stg-threshold-band      { background: #fef9c3; color: #854d0e; }
.stg-internal-gate       { background: #fed7aa; color: #9a3412; }
.stg-global-gate         { background: #fecaca; color: #991b1b; }
.badge.gate              { background: #fde68a; color: #92400e; }

.hero-meta {
  display: flex; flex-wrap: wrap; gap: 10px; margin: 8px 0 16px 0;
}
.meta-pill {
  background: var(--code-bg);
  padding: 6px 12px;
  border-radius: 999px;
  font-size: 12.5px;
}
.meta-pill strong { color: var(--accent); }

p.desc { color: var(--ink-soft); }

.audit-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(140px, 1fr));
  gap: 10px;
  margin: 8px 0 18px 0;
}
.audit-card {
  background: var(--code-bg);
  padding: 10px 12px;
  border-radius: 8px;
  text-align: center;
}
.audit-n { font-size: 22px; font-weight: 700; color: var(--accent); }
.audit-k { font-size: 11.5px; color: var(--ink-soft); text-transform: uppercase; letter-spacing: 0.5px; margin-top: 2px; }

ul.bullet { padding-left: 18px; margin: 4px 0 12px; }
ul.bullet li { margin: 4px 0; }
ul.bullet code { font-size: 11.5px; }

.hidden { display: none; }
"""

JS = """
// Tiny per-table search filter.
document.addEventListener("DOMContentLoaded", () => {
  document.querySelectorAll("input.filter").forEach((inp) => {
    const tbl = document.getElementById(inp.dataset.target);
    if (!tbl) return;
    inp.addEventListener("input", () => {
      const q = inp.value.trim().toLowerCase();
      tbl.querySelectorAll("tbody tr").forEach((tr) => {
        tr.style.display = q && !tr.textContent.toLowerCase().includes(q) ? "none" : "";
      });
    });
  });

  // Smooth-scroll for nav (account for sticky header)
  document.querySelectorAll('a[href^="#"]').forEach((a) => {
    a.addEventListener("click", (ev) => {
      const id = a.getAttribute("href").slice(1);
      const el = document.getElementById(id);
      if (!el) return;
      ev.preventDefault();
      const offset = 80;
      const top = el.getBoundingClientRect().top + window.pageYOffset - offset;
      window.scrollTo({ top, behavior: "smooth" });
      if (history.pushState) history.pushState(null, "", "#" + id);
    });
  });
});
"""


def main() -> int:
    spec = json.loads(SPEC_PATH.read_text())

    meta = spec["_meta"]
    readme = spec.get("readme", {})

    sheets_html = "\n".join([
        render_readme(meta, readme),
        render_source_files(spec),
        render_source_fields(spec),
        render_derived_fields(spec),
        render_indicators(spec),
        render_thresholds(spec),
        render_building_blocks(spec),
        render_block_composition(spec),
        render_drone_score(spec),
        render_flags(spec),
        render_script_hints(spec),
    ])

    title = f"{meta.get('title', 'drone_provenance_ppk')} v{meta.get('version', '')}"

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
<footer style="text-align:center; padding: 20px 0; color: #64748b; font-size: 12px;">
  Generated from <code>drone_provenance_ppk.json</code> v{html.escape(str(meta.get('version', '')))}
  &middot; This HTML mirrors <code>drone_provenance_ppk.xlsx</code>; the JSON is the canonical source.
</footer>
<script>{JS}</script>
</body>
</html>
"""

    OUT_PATH.write_text(out)
    print(f"wrote {OUT_PATH.relative_to(SPEC_PATH.parent.parent)}")
    print(f"   size: {OUT_PATH.stat().st_size:,} bytes")
    print(f"   sections: {len(NAV_LINKS)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
