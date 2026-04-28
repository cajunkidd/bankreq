"""Standalone launcher for the Bank Reconciliation app.

Starts the FastAPI server on localhost:8000 and opens the user's default
browser. Used both for "python run.py" during development and as the
PyInstaller entrypoint for the packaged .exe.
"""
import socket
import sys
import threading
import time
import webbrowser

import uvicorn

from app.main import app

HOST = "127.0.0.1"
DEFAULT_PORT = 8000


def find_open_port(host: str, start: int) -> int:
    port = start
    for _ in range(50):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind((host, port))
                return port
            except OSError:
                port += 1
    raise RuntimeError("Could not find an open port near " + str(start))


def open_browser_when_ready(url: str) -> None:
    # Tiny delay so the server is listening before the page loads.
    time.sleep(1.0)
    webbrowser.open(url)


def main() -> None:
    port = find_open_port(HOST, DEFAULT_PORT)
    url = f"http://{HOST}:{port}"

    print()
    print("=" * 60)
    print("  Stine Bank Reconciliation")
    print("=" * 60)
    print(f"  Open: {url}")
    print("  (Your browser should open automatically.)")
    print()
    print("  To stop: close this window or press Ctrl+C.")
    print("=" * 60)
    print()

    threading.Thread(
        target=open_browser_when_ready, args=(url,), daemon=True
    ).start()

    try:
        uvicorn.run(app, host=HOST, port=port, log_level="warning")
    except KeyboardInterrupt:
        pass
    except Exception as e:
        print(f"\nServer error: {e}")
        input("Press Enter to close...")
        sys.exit(1)


if __name__ == "__main__":
    main()
