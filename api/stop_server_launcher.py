"""
Stops the OfficeToPDF Flask server running on port 5001.
Double-click the built StopServer.exe to kill it.
"""
import urllib.request
import sys
import time
import tkinter as tk
from tkinter import messagebox


def stop_server():
    try:
        urllib.request.urlopen("http://localhost:5001/shutdown", timeout=3)
        return True
    except Exception:
        return False


def main():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    success = stop_server()
    if success:
        messagebox.showinfo("OfficeToPDF", "Server stopped successfully.")
    else:
        messagebox.showerror("OfficeToPDF", "Could not reach server on port 5001.\nIt may already be stopped.")
    root.destroy()
    sys.exit(0)


if __name__ == "__main__":
    main()
