"""
Launcher for office_to_pdf Flask server.
Double-click the built .exe to start the server on http://localhost:5001
"""
import threading
import webbrowser
import time
import sys
import os

# Ensure imports resolve when running as a frozen exe
if getattr(sys, 'frozen', False):
    sys.path.insert(0, sys._MEIPASS)

from office_to_pdf import app


def open_browser():
    time.sleep(1.5)
    webbrowser.open("http://localhost:5001")


if __name__ == "__main__":
    threading.Thread(target=open_browser, daemon=True).start()
    print("Office-to-PDF server running at http://localhost:5001")
    print("Press Ctrl+C to stop.")
    app.run(host="0.0.0.0", port=5001, debug=False, use_reloader=False)
