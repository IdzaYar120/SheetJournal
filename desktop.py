"""
Desktop launcher: runs the Flask app on a private local port and shows it in a
native window (Edge WebView2 on Windows), so the journal generator can be used
as a regular double-click program — no browser, terminal or Python needed.

Build the Windows .exe with:  pyinstaller SheetJournal.spec
Run with --browser to open the default web browser instead of a native window.
"""

import socket
import sys
import threading
import webbrowser

from werkzeug.serving import make_server

from app import app

APP_TITLE = "Генератор навчальних журналів"


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _start_server() -> tuple[str, object]:
    port = _free_port()
    server = make_server("127.0.0.1", port, app, threaded=True)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    return f"http://127.0.0.1:{port}/", server


def _run_in_browser(url: str, server) -> None:
    """Open the app in the default browser and keep a small control window
    alive so the user has a visible way to quit (the .exe has no console)."""
    webbrowser.open(url)
    try:
        import tkinter as tk
    except ImportError:
        try:
            threading.Event().wait()
        except KeyboardInterrupt:
            pass
        server.shutdown()
        return

    root = tk.Tk()
    root.title(APP_TITLE)
    root.resizable(False, False)
    tk.Label(
        root,
        text=f"Програма працює у вашому браузері:\n{url}\n\n"
             "Закрийте це вікно, щоб завершити роботу.",
        padx=24, pady=16, justify="center",
    ).pack()
    tk.Button(root, text="Відкрити в браузері", command=lambda: webbrowser.open(url)).pack(pady=(0, 16))
    root.mainloop()
    server.shutdown()


def main() -> None:
    url, server = _start_server()

    if "--browser" in sys.argv:
        _run_in_browser(url, server)
        return

    try:
        import webview
    except ImportError:
        _run_in_browser(url, server)
        return

    # Show the native "Save as" dialog for /download/<token> links.
    webview.settings["ALLOW_DOWNLOADS"] = True
    webview.create_window(APP_TITLE, url, width=1280, height=860, min_size=(900, 600))
    try:
        webview.start()
    except Exception:
        # No usable WebView backend (e.g. WebView2 runtime missing) — fall back
        # to the system browser so the program still works.
        _run_in_browser(url, server)
        return
    server.shutdown()


if __name__ == "__main__":
    main()
