"""PyInstaller entry point: boot the Streamlit app and open the browser.

The Streamlit app itself is the unmodified app.py — this file only exists to
let PyInstaller package a .exe that, when double-clicked, starts Streamlit
in-process and opens the user's browser to the local server.
"""
from __future__ import annotations

import io
import multiprocessing
import os
import socket
import sys
import threading
import time
import traceback
import urllib.error
import urllib.request
import webbrowser
from pathlib import Path


LOG_FILE = Path.home() / "BankDataViewer.log"
_log_lock = threading.Lock()
_log_fp = None  # type: ignore[var-annotated]


def _open_log() -> None:
    global _log_fp
    try:
        LOG_FILE.unlink(missing_ok=True)
    except OSError:
        pass
    try:
        _log_fp = open(LOG_FILE, "a", encoding="utf-8", buffering=1)
    except OSError:
        _log_fp = None


def _log(msg: str) -> None:
    line = f"{time.strftime('%Y-%m-%d %H:%M:%S')} {msg}\n"
    with _log_lock:
        if _log_fp is not None:
            try:
                _log_fp.write(line)
                _log_fp.flush()
            except Exception:
                pass
        try:
            sys.__stdout__.write(line)
            sys.__stdout__.flush()
        except Exception:
            pass


class _Tee:
    """File-like that fan-outs writes to multiple underlying streams."""

    def __init__(self, *streams) -> None:
        self._streams = [s for s in streams if s is not None]

    def write(self, data) -> int:
        n = 0
        for s in self._streams:
            try:
                n = s.write(data) or 0
                s.flush()
            except Exception:
                pass
        return n

    def flush(self) -> None:
        for s in self._streams:
            try:
                s.flush()
            except Exception:
                pass

    def isatty(self) -> bool:
        return False

    def fileno(self) -> int:
        raise io.UnsupportedOperation


def _install_stream_tees() -> None:
    """Capture all stdout/stderr to the log file in addition to the console
    so Streamlit's own output is visible even without a console window."""
    if _log_fp is None:
        return
    sys.stdout = _Tee(sys.__stdout__, _log_fp)
    sys.stderr = _Tee(sys.__stderr__, _log_fp)


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


def _diagnose_bundle() -> None:
    """Log enough about the frozen bundle's filesystem layout to figure out
    why Streamlit might 404. We log the streamlit static directory's contents
    and where streamlit.__file__ points."""
    base = _bundle_root()
    static = base / "streamlit" / "static"
    _log(f"Bundle streamlit/static: {static} (exists={static.exists()})")
    if static.exists():
        try:
            entries = sorted(p.name for p in static.iterdir())
            _log(f"  top-level entries ({len(entries)}): {entries[:20]}")
            index_html = static / "index.html"
            _log(
                f"  index.html: exists={index_html.exists()}, "
                f"size={index_html.stat().st_size if index_html.exists() else 'n/a'}"
            )
        except OSError as e:
            _log(f"  error listing static dir: {e}")
    try:
        import streamlit  # type: ignore

        _log(f"streamlit.__file__: {streamlit.__file__}")
        st_dir = Path(streamlit.__file__).parent
        st_static = st_dir / "static"
        _log(
            f"streamlit static (via __file__): {st_static} "
            f"(exists={st_static.exists()})"
        )
        if st_static.exists():
            idx = st_static / "index.html"
            _log(f"  index.html via __file__: exists={idx.exists()}")
    except Exception as e:  # noqa: BLE001
        _log(f"streamlit import diagnostic failed: {e}")


def _open_browser_when_ready(port: int) -> None:
    deadline = time.time() + 90
    url = f"http://localhost:{port}/"
    health_url = f"http://localhost:{port}/_stcore/health"
    last_status = ""
    next_progress = time.time() + 5
    health_logged = False
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(url, timeout=1) as resp:
                _log(f"HTTP {resp.status} from {url}; opening browser.")
                webbrowser.open(url)
                return
        except urllib.error.HTTPError as e:
            last_status = f"HTTP {e.code}"
            if not health_logged:
                try:
                    with urllib.request.urlopen(health_url, timeout=1) as hresp:
                        body = hresp.read(64).decode("utf-8", "replace")
                        _log(
                            f"Probe {health_url}: HTTP {hresp.status} body={body!r}"
                        )
                        health_logged = True
                except Exception as he:  # noqa: BLE001
                    _log(f"Probe {health_url}: {type(he).__name__}: {he}")
                    health_logged = True
        except urllib.error.URLError as e:
            last_status = f"URLError: {e.reason}"
        except Exception as e:  # noqa: BLE001
            last_status = f"{type(e).__name__}: {e}"
        if time.time() >= next_progress:
            _log(f"Still waiting for {url} ({last_status})")
            next_progress = time.time() + 5
        time.sleep(0.4)
    _log(f"Timed out after 90s waiting for {url} (last={last_status}).")


def main() -> None:
    _open_log()
    _install_stream_tees()
    _log(f"Starting Bank Data Viewer (frozen={getattr(sys, 'frozen', False)}).")
    _log(f"Python: {sys.version.split()[0]}, executable: {sys.executable}")

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

    flag_options = {
        "browser.gatherUsageStats": False,
        "browser.serverAddress": "localhost",
        "browser.serverPort": port,
        "server.headless": True,
        "server.port": port,
        "server.address": "127.0.0.1",
        "server.fileWatcherType": "none",
        "server.runOnSave": False,
        "global.developmentMode": False,
    }

    _diagnose_bundle()

    threading.Thread(
        target=_open_browser_when_ready, args=(port,), daemon=True
    ).start()

    try:
        _log("Importing streamlit.web.bootstrap ...")
        from streamlit.web import bootstrap

        _log(f"Calling streamlit bootstrap.run with flag_options={flag_options}")
        bootstrap.run(str(app_path), False, [], flag_options)
        _log("bootstrap.run returned (server stopped).")
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
