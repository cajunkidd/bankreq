"""PyInstaller entry point: boot the Streamlit app and open the browser.

The Streamlit app itself is the unmodified app.py — this file only exists to
let PyInstaller package a .exe that, when double-clicked, starts Streamlit
in-process and opens the user's browser to the local server.
"""
from __future__ import annotations

import multiprocessing
import os
import socket
import sys
import threading
import time
import webbrowser
from pathlib import Path


def _bundle_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    return Path(__file__).resolve().parent


def _find_free_port(preferred: int = 8501) -> int:
    for port in (preferred, 8502, 8503, 0):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return s.getsockname()[1]
            except OSError:
                continue
    return preferred


def _silence_missing_streams() -> None:
    # PyInstaller --windowed sets stdout/stderr to None; tornado/streamlit
    # crash when they try to write. Redirect to devnull.
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w")


def main() -> None:
    _silence_missing_streams()
    base = _bundle_root()
    app_path = base / "app.py"
    port = _find_free_port(8501)

    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
    os.environ["STREAMLIT_SERVER_PORT"] = str(port)
    os.environ["STREAMLIT_SERVER_ADDRESS"] = "127.0.0.1"
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"

    def open_browser_when_ready() -> None:
        deadline = time.time() + 30
        url = f"http://localhost:{port}"
        while time.time() < deadline:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.settimeout(0.5)
                try:
                    s.connect(("127.0.0.1", port))
                    webbrowser.open(url)
                    return
                except OSError:
                    time.sleep(0.3)

    threading.Thread(target=open_browser_when_ready, daemon=True).start()

    from streamlit.web import bootstrap

    bootstrap.run(str(app_path), False, [], {})


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
