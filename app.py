#!/usr/bin/env python3
"""
CBMI Loop — desktop entry point.

Starts the FastAPI server on a free localhost port and opens the UI in a
native window (pywebview) or, if pywebview isn't available, the default
browser. Closing the window shuts the server down and terminates any pipeline
steps still running.

Dev usage:
    python3 app.py               # window (or browser) + server
    python3 app.py --serve-only  # just the server, prints the URL
"""
from __future__ import annotations

import argparse
import multiprocessing
import socket
import sys
import threading
import time
import webbrowser


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _wait_ready(port: int, timeout: float = 30.0) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.5):
                return True
        except OSError:
            time.sleep(0.15)
    return False


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--serve-only", action="store_true",
                    help="run the server without opening a window")
    ap.add_argument("--port", type=int, default=0)
    args = ap.parse_args()

    from cbmi import APP_NAME, APP_VERSION, appdata, build_info, runner

    appdata.materialize()

    ok, note = build_info.check(appdata.app_data_dir())
    if not ok:
        print(f"{APP_NAME}: {note}", file=sys.stderr)
        try:
            import webview  # noqa: F401
            import webview as _wv
            _wv.create_window(APP_NAME, html=f"<h2 style='font-family:sans-serif'>"
                                             f"{APP_NAME}</h2><p>{note}</p>",
                              width=520, height=220)
            _wv.start()
        except Exception:  # noqa: BLE001
            pass
        return 3

    import uvicorn

    from server.main import app as fastapi_app

    port = args.port or _free_port()
    url = f"http://127.0.0.1:{port}/"
    log_file = appdata.logs_dir() / "server.log"

    config = uvicorn.Config(fastapi_app, host="127.0.0.1", port=port,
                            log_level="warning")
    server = uvicorn.Server(config)
    t = threading.Thread(target=server.run, daemon=True)
    t.start()

    if not _wait_ready(port):
        print(f"{APP_NAME}: server failed to start (see {log_file})",
              file=sys.stderr)
        return 1

    print(f"{APP_NAME} {APP_VERSION} — {url}" + (f"  [{note}]" if note else ""))

    def _shutdown() -> None:
        runner.terminate_all()
        server.should_exit = True

    if args.serve_only:
        try:
            while t.is_alive():
                time.sleep(0.5)
        except KeyboardInterrupt:
            _shutdown()
        return 0

    try:
        import webview
        window = webview.create_window(APP_NAME, url, width=1440, height=920,
                                       min_size=(1024, 700))
        del window
        webview.start()          # blocks until the window closes
        _shutdown()
    except ImportError:
        webbrowser.open(url)
        print("(pywebview not installed — opened in the default browser; "
              "Ctrl+C to quit)")
        try:
            while t.is_alive():
                time.sleep(0.5)
        except KeyboardInterrupt:
            _shutdown()
    return 0


if __name__ == "__main__":
    multiprocessing.freeze_support()   # required for frozen Windows builds
    raise SystemExit(main())
