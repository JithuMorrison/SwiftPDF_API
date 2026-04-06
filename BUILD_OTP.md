# Building OfficeToPDF.exe

## Files involved

```
api/
  office_to_pdf.py          ← the actual Flask API (Word/Excel/PPT → PDF)
  office_to_pdf_launcher.py ← entry point for OfficeToPDF.exe
  stop_server_launcher.py   ← entry point for StopServer.exe
OfficeToPDF.spec            ← PyInstaller build config for the server exe
StopServer.spec             ← PyInstaller build config for the stop exe
dist/
  OfficeToPDF.exe           ← starts the server
  StopServer.exe            ← stops the server
```

---

## Why office_to_pdf_launcher.py exists

PyInstaller needs a single entry-point script — a file with a `if __name__ == "__main__"` block that starts everything.

`office_to_pdf.py` is a Flask API module. It defines routes and a Flask `app` object, but it doesn't start the server on its own when bundled as an exe. If you pointed PyInstaller directly at it, double-clicking the exe would do nothing visible — the server would start and immediately exit with no feedback.

`office_to_pdf_launcher.py` solves this by:

1. Importing the `app` from `office_to_pdf.py`
2. Starting the Flask server with `app.run()`
3. Opening `http://localhost:5001` in the browser automatically after 1.5 seconds
4. Printing a status message so the user knows it's running

This separation keeps the API code clean and reusable while giving PyInstaller a proper standalone entry point.

---

## Prerequisites

- Python 3.10+
- Microsoft Office installed (Word/Excel/PowerPoint) for best PDF quality
- LibreOffice optional — used as fallback if MS Office COM fails
- No MS Office or LibreOffice? The basic python fallback still works

```cmd
python -m venv venv
venv\Scripts\pip install flask flask-cors pywin32 python-docx reportlab python-pptx pandas fpdf openpyxl Pillow pyinstaller
```

---

## Build OfficeToPDF.exe (starts the server)

### Option 1 — using the .spec file (recommended)

```cmd
venv\Scripts\pyinstaller OfficeToPDF.spec
```

### Option 2 — from scratch

```cmd
venv\Scripts\pyinstaller --onefile --noconsole --name OfficeToPDF api\office_to_pdf_launcher.py --paths api
```

Output: `dist\OfficeToPDF.exe`

---

## Build StopServer.exe (stops the server)

### Option 1 — using the .spec file (recommended)

```cmd
venv\Scripts\pyinstaller StopServer.spec
```

### Option 2 — from scratch

```cmd
venv\Scripts\pyinstaller --onefile --noconsole --name StopServer api\stop_server_launcher.py
```

Output: `dist\StopServer.exe`

---

## Pre-built downloads (Windows x64)

Don't want to build? Download the pre-built executables directly:

| File                                  | Download                                                                                                         |
| ------------------------------------- | ---------------------------------------------------------------------------------------------------------------- |
| `OfficeToPDF.exe` — starts the server | [Download from Google Drive](https://drive.google.com/file/d/12EftoV1bFdw2OQdnRiwPMAb8Fptqg2bM/view?usp=sharing) |
| `StopServer.exe` — stops the server   | [Download from Google Drive](https://drive.google.com/file/d/1vNnWQLZMf0PMOY1KfuRdzZQpKxShDVv7/view?usp=sharing) |

Place both in the same folder and use them together.

---

## What you can delete after building

| Path                            | Safe to delete?                               |
| ------------------------------- | --------------------------------------------- |
| `build\`                        | Yes — intermediate files only                 |
| `dist\OfficeToPDF.exe`          | No — this is the output                       |
| `dist\StopServer.exe`           | No — this is the output                       |
| `OfficeToPDF.spec`              | No — needed to rebuild without retyping flags |
| `StopServer.spec`               | No — needed to rebuild without retyping flags |
| `api\office_to_pdf_launcher.py` | No — required by OfficeToPDF.spec             |
| `api\stop_server_launcher.py`   | No — required by StopServer.spec              |

---

## Runtime requirements on the target machine

The exe is self-contained (no Python needed), but conversion quality depends on what's installed:

- Microsoft Office installed → uses Word/Excel/PowerPoint COM for best fidelity
- LibreOffice installed → used as fallback if MS Office is not present
- Neither installed → basic python fallback kicks in (simpler output, always works)
