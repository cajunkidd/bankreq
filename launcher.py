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
import traceback
import webbrowser
from pathlib import Path


LOG_FILE = Path.home() / "BankDataViewer.log"


def _log(msg: str) -> None:
    line = f"{time.strftime('%Y-%m-%d %H:%M:%S')} {msg}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass
    try:
        print(line, flush=True)
    except Exception:
        pass


def _bundle_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    return Path(__file__).resolve().parent


def _find_free_port(preferred: int = 8501) -> int:
    for port in (preferred, 8502, 8503, 8504, 0):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return s.getsockname()[1]
            except OSError:
                continue
    return preferred


def _silence_missing_streams() -> None:
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w")


def _open_browser_when_ready(port: int) -> None:
    deadline = time.time() + 60
    url = f"http://localhost:{port}"
    while time.time() < deadline:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.5)
            try:
                s.connect(("127.0.0.1", port))
                _log(f"Server up on {url}; opening browser.")
                webbrowser.open(url)
                return
            except OSError:
                time.sleep(0.3)
    _log(f"Timed out waiting for server on {url}.")


def main() -> None:
    _silence_missing_streams()
    try:
        LOG_FILE.unlink(missing_ok=True)
    except OSError:
        pass
    _log(f"Starting Bank Data Viewer (frozen={getattr(sys, 'frozen', False)}).")

    base = _bundle_root()
    app_path = base / "app.py"
    _log(f"Bundle root: {base}")
    _log(f"App path:    {app_path} (exists={app_path.exists()})")
    if not app_path.exists():
        _log("ERROR: app.py missing from bundle.")
        time.sleep(8)
        return

    port = _find_free_port(8501)
    _log(f"Bound port:  {port}")

    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
    os.environ["STREAMLIT_SERVER_PORT"] = str(port)
    os.environ["STREAMLIT_SERVER_ADDRESS"] = "127.0.0.1"
    os.environ["STREAMLIT_SERVER_FILE_WATCHER_TYPE"] = "none"
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"

    threading.Thread(target=_open_browser_when_ready, args=(port,), daemon=True).start()

    try:
        from streamlit.web import bootstrap

        _log("Calling streamlit bootstrap.run...")
        bootstrap.run(str(app_path), False, [], {})
    except SystemExit:
        raise
    except BaseException:
        tb = traceback.format_exc()
        _log("FATAL: Streamlit bootstrap raised an exception:\n" + tb)
        print(
            "\nThe app failed to start. A log has been written to:\n"
            f"  {LOG_FILE}\n\nPress Enter to close this window."
        )
        try:
            input()
        except EOFError:
            pass


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
