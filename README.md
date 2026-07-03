# CBMI Loop — Desktop App

Self-contained desktop packaging of the CBMI Loop app (Load page + 4 hardware
pages + the four scoring pipelines). This folder is the future repo root for
CI builds (.exe / .dmg).

## Layout

| Path | What | Editable? |
|---|---|---|
| `app.py` | Entry point: server + native window | yes |
| `cbmi/` | Packaging layer: app-data, step runner, build identity | yes |
| `server/` | FastAPI app (desktop edition of Frontend/server) | yes |
| `webui/` | Load + hardware pages (copy of Frontend/app) | yes |
| `vendor.py` | Generates the two below from `../*_CodeBase` | yes |
| `cbmi_pipelines/` | GENERATED — vendored pipeline code | **no — re-vendor** |
| `runtime_data/` | GENERATED — pipeline data (specs, libs, caches) | **no — re-vendor** |

The four source `*_CodeBase` folders are **never modified**. When upstream
pipeline code changes, run `python3 vendor.py` to refresh.

## Run (dev)

```bash
python3 -m pip install --user -r requirements.txt
python3 app.py                 # native window (or browser fallback)
python3 app.py --serve-only    # server only, prints URL
```

Smoke-test a pipeline against the source repo's sample data:

```bash
python3 -m server.orchestrator --subsystem checkpoint --sample
```

## Where user data lives

Everything written at runtime goes to per-user app-data (the install dir stays
read-only):

- macOS: `~/Library/Application Support/CBMI Loop/`
- Windows: `%LOCALAPPDATA%\CBMI Loop\`

Inside: `workspace/` (pipeline roots + staged inputs), `results/` (latest
per-subsystem results — restored on launch), `runs/` (history), `logs/`.

## How the pipelines run without python3

`cbmi/runner.py` spawns children of the app's own process (multiprocessing
"spawn"); each child imports the vendored package
(`cbmi_pipelines.<key>.scripts.<step>`) and replays the exact `main(...)` call
the script's own `__main__` block makes (see generated
`cbmi_pipelines/entrypoints.py`). This works identically in dev and inside a
frozen (Nuitka/PyInstaller) binary — which is what allows shipping the
pipeline code compiled, with no `.py` source on disk.
