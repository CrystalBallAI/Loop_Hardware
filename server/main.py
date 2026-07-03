"""
main.py — FastAPI app, desktop edition.

Routes (same as the dev server, plus run history):
  GET  /                      -> Load page (webui/load.html + load-integration.*)
  GET  /hw                    -> Hardware SPA (webui/hw.html + hw-integration.*)
  GET  /static/*              -> integration js/css
  GET  /api/subsystems        -> registry summary
  POST /api/run               -> multipart upload + start background run -> {runId}
  GET  /api/run/{runId}       -> live per-subsystem status
  GET  /api/results/{sub}     -> hardware-page result object (latest)
  GET  /api/download/{sub}/xlsx | /provenance
  GET  /api/runs              -> run history (persisted across restarts)
  GET  /api/runs/{runId}      -> one historical run
  POST /api/runs/{runId}/activate -> make that run's results the current ones
  GET  /api/runs/{runId}/download/{sub}/{kind}  (kind: xlsx | provenance)
  GET  /api/app               -> version / tester / expiry note (for the UI)
"""
from __future__ import annotations

import json
import shutil
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.responses import (FileResponse, HTMLResponse, JSONResponse,
                               PlainTextResponse)
from fastapi.staticfiles import StaticFiles

from cbmi import APP_VERSION, appdata, build_info

from . import orchestrator, run_store
from .subsystems import APP_DIR, SUBSYSTEMS, get

app = FastAPI(title="CBMI Loop")

STATIC_DIR = APP_DIR / "static"
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

# ---- HTML pages (integration tags injected once, cached) ------------------
_PAGE_CACHE: dict = {}


def _ver(name: str) -> int:
    try:
        return int((STATIC_DIR / name).stat().st_mtime)
    except OSError:
        return 0


def _serve_page(filename: str, css: str, js: str) -> HTMLResponse:
    try:
        pv = int((APP_DIR / filename).stat().st_mtime)
    except OSError:
        pv = 0
    cv, jv = _ver(css), _ver(js)
    key = (filename, pv, cv, jv)
    if key not in _PAGE_CACHE:
        html = (APP_DIR / filename).read_text(encoding="utf-8", errors="ignore")
        inject = (f'\n<link rel="stylesheet" href="/static/{css}?v={cv}">\n'
                  f'<script src="/static/{js}?v={jv}"></script>\n')
        if "</body>" in html:
            html = html.replace("</body>", inject + "</body>", 1)
        else:
            html += inject
        _PAGE_CACHE.clear()
        _PAGE_CACHE[key] = html
    return HTMLResponse(_PAGE_CACHE[key])


@app.get("/", response_class=HTMLResponse)
def page_load():
    return _serve_page("load.html", "load-integration.css", "load-integration.js")


@app.get("/hw", response_class=HTMLResponse)
def page_hw():
    return _serve_page("hw.html", "hw-integration.css", "hw-integration.js")


# ---- API ------------------------------------------------------------------
@app.get("/api/subsystems")
def api_subsystems():
    return {
        "subsystems": [
            {
                "key": s.key, "label": s.label, "module": s.module,
                "perPoint": s.per_point,
                "inputs": [sl.input_id for sl in s.slots],
                "pointSlots": [ps.slot_id for ps in s.point_slots],
            }
            for s in SUBSYSTEMS.values()
        ]
    }


@app.get("/api/app")
def api_app():
    ok, note = build_info.check(appdata.app_data_dir())
    return {"version": APP_VERSION, "tester": build_info.TESTER_ID,
            "ok": ok, "note": note}


@app.post("/api/run")
async def api_run(request: Request):
    form = await request.form()

    def _list(field):
        try:
            return list(json.loads(form.get(field) or "[]"))
        except ValueError:
            return []

    explicit_subs = _list("subsystems")

    files_by_sub: dict = {}
    for field, value in form.multi_items():
        if not hasattr(value, "filename"):
            continue
        key = field.split("__", 1)[0]
        if key not in SUBSYSTEMS:
            continue
        files_by_sub.setdefault(key, []).append((field, value))

    run_subs = [k for k in (explicit_subs or sorted(files_by_sub)) if k in SUBSYSTEMS]
    if not run_subs:
        return JSONResponse({"error": "No subsystems to run (no files uploaded)."},
                            status_code=400)

    staged_errors = []
    for key in run_subs:
        sub = SUBSYSTEMS[key]
        orchestrator.prepare_subsystem(sub)
        for field, uf in files_by_sub.get(key, []):
            dest = orchestrator.destination_for(sub, field, uf.filename or "file")
            if dest is None:
                staged_errors.append(f"{key}: unrecognized field {field}")
                continue
            target, also = dest
            target.parent.mkdir(parents=True, exist_ok=True)
            with target.open("wb") as out:
                shutil.copyfileobj(uf.file, out)
            if also:
                also.parent.mkdir(parents=True, exist_ok=True)
                shutil.copyfile(target, also)

    run_id = orchestrator.start_run(run_subs)
    return {"runId": run_id, "subsystems": run_subs, "stagedWarnings": staged_errors}


@app.get("/api/run/{run_id}")
def api_run_status(run_id: str):
    run = run_store.get_run(run_id)
    if run is None:
        return JSONResponse({"error": "unknown run"}, status_code=404)
    return run


@app.get("/api/results/{sub}")
def api_results(sub: str):
    try:
        key = get(sub).key
    except KeyError:
        return JSONResponse({"error": "unknown subsystem"}, status_code=404)
    result = orchestrator.load_result(key)
    if result is None:
        return JSONResponse({"error": "no result yet"}, status_code=404)
    return result


@app.get("/api/download/{sub}/xlsx")
def api_xlsx(sub: str):
    try:
        s = get(sub)
    except KeyError:
        return PlainTextResponse("unknown subsystem", status_code=404)
    path = s.codebase / s.xlsx_name
    if not path.exists():
        return PlainTextResponse("Excel not generated yet", status_code=404)
    return FileResponse(str(path),
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        filename=path.name)


@app.get("/api/download/{sub}/provenance")
def api_provenance(sub: str):
    try:
        s = get(sub)
    except KeyError:
        return PlainTextResponse("unknown subsystem", status_code=404)
    prov = orchestrator._provenance_file(s)
    if not prov:
        return PlainTextResponse("Provenance not generated yet", status_code=404)
    return FileResponse(str(prov), media_type="text/html", filename=prov.name)


# ---- run history ------------------------------------------------------------
@app.get("/api/runs")
def api_runs():
    return {"runs": orchestrator.list_runs()}


@app.get("/api/runs/{run_id}")
def api_run_history(run_id: str):
    run = orchestrator.get_run_history(run_id)
    if run is None:
        return JSONResponse({"error": "unknown run"}, status_code=404)
    return run


@app.post("/api/runs/{run_id}/activate")
def api_run_activate(run_id: str):
    activated = orchestrator.activate_run(run_id)
    if not activated:
        return JSONResponse({"error": "run has no stored results"}, status_code=404)
    return {"activated": activated}


@app.get("/api/runs/{run_id}/download/{sub}/{kind}")
def api_run_download(run_id: str, sub: str, kind: str):
    try:
        s = get(sub)
    except KeyError:
        return PlainTextResponse("unknown subsystem", status_code=404)
    sdir = appdata.runs_dir() / run_id / s.key
    if kind == "xlsx":
        path = sdir / Path(s.xlsx_name).name
        media = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif kind == "provenance":
        cands = sorted(sdir.glob("*provenance*.html"))
        path = cands[-1] if cands else sdir / "missing"
        media = "text/html"
    else:
        return PlainTextResponse("unknown kind", status_code=404)
    if not path.exists():
        return PlainTextResponse("not stored for this run", status_code=404)
    return FileResponse(str(path), media_type=media, filename=path.name)


@app.get("/api/diagnostics")
def api_diagnostics():
    """Offline diagnostics bundle the tester can email us: logs + versions +
    run-history index. Contains NO scoring IP (specs/libraries stay encrypted
    and are not included) and no uploaded input files."""
    import io
    import platform
    import sys
    import time
    import zipfile

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        info = {
            "app": "CBMI Loop",
            "version": APP_VERSION,
            "tester": build_info.TESTER_ID,
            "python": sys.version,
            "platform": platform.platform(),
            "generatedAt": int(time.time()),
        }
        z.writestr("info.json", json.dumps(info, indent=2))
        z.writestr("runs.json", json.dumps(orchestrator.list_runs(), indent=2))
        logs = appdata.logs_dir()
        if logs.exists():
            for p in sorted(logs.glob("*.log")):
                try:
                    z.writestr(f"logs/{p.name}", p.read_text(encoding="utf-8",
                                                             errors="replace"))
                except OSError:
                    pass
    buf.seek(0)
    from fastapi.responses import Response
    return Response(
        buf.getvalue(), media_type="application/zip",
        headers={"Content-Disposition": 'attachment; filename="cbmi-diagnostics.zip"'})


@app.get("/api/health")
def health():
    return {"ok": True, "version": APP_VERSION, "subsystems": list(SUBSYSTEMS)}
