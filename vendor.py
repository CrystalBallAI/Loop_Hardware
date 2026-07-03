#!/usr/bin/env python3
"""
vendor.py — generate Desktop_App/cbmi_pipelines/ + runtime_data/ from the four
source *_CodeBase folders. THE SOURCES ARE NEVER MODIFIED; re-run this script
any time upstream changes to refresh the vendored copies.

What it does per subsystem:
  1. Copies scripts/*.py (+ scripts/parsers/*.py) into
     cbmi_pipelines/<key>/scripts[/parsers]/ as a proper package.
  2. Rewrites the codebases' flat sys.path-based imports (`import common`,
     `import parse_rinex`, `from stage2_merge import x`) into unique package
     imports (`from cbmi_pipelines.<key>.scripts import common`). This is what
     lets all four pipelines coexist inside ONE compiled binary — their module
     names collide otherwise (each has its own `common`, `parse_rinex`, ...).
     The original sys.path.insert lines are left in place: they are harmless
     no-ops once no flat import remains.
  3. Patches build_provenance_html.py (the only script that resolves paths
     from __file__ instead of the paths.json location) to honor the
     CBMI_PIPELINE_ROOT env var set by the app's step runner.
  4. Copies the runtime DATA each pipeline reads relative to its root
     (paths.json, spec JSON, recommendations library, offline API caches)
     into runtime_data/<key>/ — at app startup these are materialized into
     the user's writable app-data workspace.
  5. Verifies: no flat local imports remain, every file byte-compiles, and
     every step module imports cleanly.

Usage:  python3 vendor.py [--skip-import-check]
"""
from __future__ import annotations

import argparse
import py_compile
import re
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent          # Desktop_App/
REPO = HERE.parent                              # Loop_CodeBase/
PKG_ROOT = HERE / "cbmi_pipelines"
DATA_ROOT = HERE / "runtime_data"

# Files under scripts/ that are dev-only and must NOT ship.
EXCLUDE_SCRIPTS = {"test_scenarios.py", "make_sample_data.py"}

# data_files: (relative path, is_ip). is_ip files are AES-encrypted by --encrypt
# (the spec JSON + the recommendation library — the actual scoring IP). paths.json
# is plain config, never encrypted.
PIPELINES = {
    "drone": {
        "src": "Drone_CodeBase",
        "data_files": [
            ("paths.json", False),
            ("drone_provenance_ppk/drone_provenance_ppk.json", True),
            ("Drone_Recommendations/drone_indicator_library_v2_1.json", True),
        ],
        "data_dirs": ["cache/openmeteo"],
        "mkdirs": ["outputs"],
    },
    "basestation": {
        "src": "BaseStation_CodeBase",
        "data_files": [
            ("paths.json", False),
            ("base_station_confidence_score/base_station_confidence_score.json", True),
            ("BaseStation_Recommendations/base_station_indicator_library_v2_1.json", True),
        ],
        "data_dirs": ["cache/noaa_swpc"],
        # sample_data/: stage1 consults <root>/sample_data/hardware.json for the
        # optional hardware override (the Load page stages it there).
        "mkdirs": ["outputs", "sample_data"],
    },
    "gcp": {
        "src": "GCP_CodeBase",
        "data_files": [
            ("paths.json", False),
            ("gcp_confidence_score/gcp_confidence_score.json", True),
            ("GCP_Recommendations/gcp_indicator_library_v2_1.json", True),
        ],
        "data_dirs": ["cache/noaa_swpc"],
        "mkdirs": ["outputs"],
    },
    "checkpoint": {
        "src": "CheckPoint_CodeBase",
        "data_files": [
            ("paths.json", False),
            ("check_point_confidence_score/check_point_confidence_score.json", True),
            ("CheckPoint_Recommendations/check_point_indicator_library_v2_1.json", True),
        ],
        "data_dirs": ["cache/noaa_swpc"],
        "mkdirs": ["outputs"],
    },
}

IMPORT_RE = re.compile(r"^(\s*)import ([A-Za-z_][A-Za-z0-9_]*)(\s*(?:#.*)?)$")
FROM_RE = re.compile(r"^(\s*)from ([A-Za-z_][A-Za-z0-9_]*) import (.+)$")
MAIN_CALL_RE = re.compile(r"main\((?:sys\.argv)?\)")
STEP_SCRIPTS = ("run_pipeline", "compute_recommendations",
                "export_results_xlsx", "build_provenance_html")
PROV_FILE_EXPR = "Path(__file__).resolve().parent.parent"
# parenthesized: drone uses this expression INLINE (`<expr> / "dir" / "file"`),
# so the replacement must bind tighter than the / operator
PROV_ENV_EXPR = ('(Path(__os.environ["CBMI_PIPELINE_ROOT"]) '
                 'if __os.environ.get("CBMI_PIPELINE_ROOT") '
                 'else Path(__file__).resolve().parent.parent)')


def rewrite_imports(text: str, key: str, stem_map: dict) -> tuple[str, int]:
    """Rewrite flat local imports to unique package imports. Line-based so we
    never touch strings/docstrings spanning lines (import statements at any
    indent are single lines in these codebases — verified by audit)."""
    out, n = [], 0
    for line in text.splitlines(keepends=False):
        m = IMPORT_RE.match(line)
        if m and m.group(2) in stem_map:
            indent, stem, trail = m.group(1), m.group(2), m.group(3) or ""
            out.append(f"{indent}from cbmi_pipelines.{key}.{stem_map[stem]} "
                       f"import {stem}{trail}")
            n += 1
            continue
        m = FROM_RE.match(line)
        if m and m.group(2) in stem_map:
            indent, stem, names = m.groups()
            out.append(f"{indent}from cbmi_pipelines.{key}.{stem_map[stem]}.{stem} "
                       f"import {names}")
            n += 1
            continue
        out.append(line)
    return "\n".join(out) + "\n", n


def patch_provenance(text: str, path: Path) -> str:
    """build_provenance_html.py resolves ROOT/SPEC_PATH/OUT_PATH from __file__
    (the only script that does). Point it at CBMI_PIPELINE_ROOT instead, keeping
    the __file__ form as dev fallback."""
    if PROV_FILE_EXPR not in text:
        raise SystemExit(f"vendor: provenance pattern not found in {path} — "
                         f"upstream layout changed, update vendor.py")
    text = text.replace(PROV_FILE_EXPR, PROV_ENV_EXPR)
    # os import under a private alias so we can't collide with script names
    text = text.replace("import html\n", "import html\nimport os as __os\n", 1)
    if "import os as __os" not in text:
        raise SystemExit(f"vendor: could not insert os import in {path}")
    return text


def extract_entrypoint(key: str, script: str, text: str, path: Path) -> str:
    """Record how the script's own __main__ block calls main() — the runner
    replays exactly that (with sys.argv set), so per-script argv conventions
    (main() vs main(sys.argv)) never need guessing."""
    tail = text[text.find('__name__'):]
    m = MAIN_CALL_RE.search(tail)
    if not m:
        raise SystemExit(f"vendor: no main() call found in __main__ of {path}")
    return m.group(0)


def vendor_one(key: str, cfg: dict) -> dict:
    src = REPO / cfg["src"]
    if not src.exists():
        raise SystemExit(f"vendor: source codebase missing: {src}")

    pkg = PKG_ROOT / key
    if pkg.exists():
        shutil.rmtree(pkg)
    (pkg / "scripts" / "parsers").mkdir(parents=True)

    # ---- collect script files + build the stem -> subpackage map ----------
    scripts = sorted(p for p in (src / "scripts").glob("*.py")
                     if p.name not in EXCLUDE_SCRIPTS)
    parsers = sorted((src / "scripts" / "parsers").glob("*.py"))
    stem_map = {p.stem: "scripts" for p in scripts}
    for p in parsers:
        if p.stem in stem_map:
            raise SystemExit(f"vendor: stem collision scripts/parsers: {p.stem}")
        stem_map[p.stem] = "scripts.parsers"

    stats = {"files": 0, "rewrites": 0, "entrypoints": {}}
    for group, dest in ((scripts, pkg / "scripts"),
                        (parsers, pkg / "scripts" / "parsers")):
        for p in group:
            text = p.read_text(encoding="utf-8")
            text, n = rewrite_imports(text, key, stem_map)
            if p.name == "build_provenance_html.py":
                text = patch_provenance(text, p)
            if p.stem in STEP_SCRIPTS:
                stats["entrypoints"][f"{key}.{p.stem}"] = \
                    extract_entrypoint(key, p.stem, text, p)
            (dest / p.name).write_text(text, encoding="utf-8")
            stats["files"] += 1
            stats["rewrites"] += n

    for init in (pkg, pkg / "scripts", pkg / "scripts" / "parsers"):
        (init / "__init__.py").write_text("", encoding="utf-8")

    # ---- runtime data ------------------------------------------------------
    droot = DATA_ROOT / key
    if droot.exists():
        shutil.rmtree(droot)
    for rel, _is_ip in cfg["data_files"]:
        s, d = src / rel, droot / rel
        if not s.exists():
            raise SystemExit(f"vendor: required data file missing: {s}")
        d.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(s, d)
    for rel in cfg["data_dirs"]:
        s, d = src / rel, droot / rel
        if s.exists():
            shutil.copytree(s, d)
        else:
            d.mkdir(parents=True, exist_ok=True)
    for rel in cfg["mkdirs"]:
        (droot / rel).mkdir(parents=True, exist_ok=True)
        (droot / rel / ".keep").write_text("", encoding="utf-8")

    # ---- verify: no flat local import survived + everything compiles ------
    leftovers = []
    for p in pkg.rglob("*.py"):
        for i, line in enumerate(p.read_text(encoding="utf-8").splitlines(), 1):
            m = IMPORT_RE.match(line) or FROM_RE.match(line)
            if m and m.group(2) in stem_map:
                leftovers.append(f"{p}:{i}: {line.strip()}")
        py_compile.compile(str(p), doraise=True)
    if leftovers:
        raise SystemExit("vendor: flat local imports survived rewrite:\n  "
                         + "\n  ".join(leftovers))
    return stats


def verify_imports() -> None:
    """Import every step module in a clean child interpreter — proves the
    rewritten package graph is self-contained (what a frozen build relies on)."""
    steps = ["run_pipeline", "compute_recommendations",
             "export_results_xlsx", "build_provenance_html"]
    mods = [f"cbmi_pipelines.{k}.scripts.{s}" for k in PIPELINES for s in steps]
    code = ("import importlib\n"
            + "\n".join(f"importlib.import_module({m!r})" for m in mods)
            + "\nprint('all', len(%d), 'step modules import cleanly')" % len(mods))
    code = code.replace("len(%d)" % len(mods), str(len(mods)))
    r = subprocess.run([sys.executable, "-c", code], cwd=str(HERE),
                       capture_output=True, text=True)
    if r.returncode != 0:
        raise SystemExit(f"vendor: import verification failed:\n{r.stderr[-2000:]}")
    print(r.stdout.strip())


def encrypt_ip() -> int:
    """AES-encrypt the spec + recommendation-library files in runtime_data.
    Idempotent (skips already-encrypted). Run after vendoring."""
    from cbmi import crypt
    n = 0
    for key, cfg in PIPELINES.items():
        for rel, is_ip in cfg["data_files"]:
            if not is_ip:
                continue
            p = DATA_ROOT / key / rel
            if crypt.encrypt_file(p):
                n += 1
                print(f"  encrypted {key}/{rel}")
            else:
                print(f"  already-encrypted {key}/{rel}")
    return n


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--skip-import-check", action="store_true")
    ap.add_argument("--encrypt", action="store_true",
                    help="AES-encrypt the spec/library IP files after vendoring")
    ap.add_argument("--encrypt-only", action="store_true",
                    help="ONLY (idempotently) encrypt the already-committed runtime_data IP; "
                         "do NOT re-vendor. Use on build/CI machines that don't have the "
                         "source *_CodeBase folders (the vendored artifacts are committed).")
    args = ap.parse_args()

    if args.encrypt_only:
        print("encrypt-only: ensuring committed runtime_data IP is encrypted "
              "(no re-vendor)")
        got = encrypt_ip()
        print(f"encrypted {got} file(s); {8 - got} already-encrypted")
        print("vendor: OK (encrypt-only)")
        return 0

    PKG_ROOT.mkdir(exist_ok=True)
    (PKG_ROOT / "__init__.py").write_text("", encoding="utf-8")
    DATA_ROOT.mkdir(exist_ok=True)

    entrypoints = {}
    for key, cfg in PIPELINES.items():
        st = vendor_one(key, cfg)
        entrypoints.update(st["entrypoints"])
        print(f"vendored {key:12s} files={st['files']:3d} "
              f"import-rewrites={st['rewrites']:3d}")

    # generated dispatch table: how each step module's own __main__ calls main()
    ep = PKG_ROOT / "entrypoints.py"
    lines = ["# GENERATED by vendor.py — do not edit.",
             "# '<key>.<script>': the exact call the script's __main__ makes.",
             "ENTRY = {"]
    for k in sorted(entrypoints):
        lines.append(f"    {k!r}: {entrypoints[k]!r},")
    lines.append("}")
    ep.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"wrote {ep.relative_to(HERE)} ({len(entrypoints)} entrypoints)")
    if not args.skip_import_check:
        verify_imports()
    if args.encrypt:
        print("encrypting IP data files:")
        got = encrypt_ip()
        print(f"encrypted {got} file(s)")
    print("vendor: OK" + (" (encrypted)" if args.encrypt else " (plaintext IP)"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
