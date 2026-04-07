from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import time
import tempfile
import subprocess
import platform
import threading


app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = tempfile.mkdtemp()


def safe_remove(file_path):
    """Safely remove a file with retries."""
    for _ in range(5):
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
            break
        except PermissionError:
            time.sleep(0.5)


def convert_with_libreoffice(input_path, output_dir):
    """Convert office file to PDF using LibreOffice (cross-platform)."""
    libreoffice_cmd = None

    if platform.system() == "Windows":
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for c in candidates:
            if os.path.exists(c):
                libreoffice_cmd = c
                break
    elif platform.system() == "Darwin":
        libreoffice_cmd = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    else:
        libreoffice_cmd = "libreoffice"

    if libreoffice_cmd is None:
        raise RuntimeError("LibreOffice not found. Please install LibreOffice.")

    result = subprocess.run(
        [libreoffice_cmd, "--headless", "--convert-to", "pdf", "--outdir", output_dir, input_path],
        capture_output=True,
        text=True,
        timeout=60,
    )

    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")

    base_name = os.path.splitext(os.path.basename(input_path))[0]
    output_pdf = os.path.join(output_dir, base_name + ".pdf")

    if not os.path.exists(output_pdf):
        raise RuntimeError("PDF output not found after conversion.")

    return output_pdf


def convert_with_win32com_word(input_path, output_path):
    """Convert Word to PDF using Microsoft Word COM (Windows only)."""
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(input_path))
        doc.SaveAs(os.path.abspath(output_path), FileFormat=17)  # 17 = wdFormatPDF
        doc.Close()
    finally:
        word.Quit()


def convert_with_win32com_excel(input_path, output_path):
    """Convert Excel to PDF using Microsoft Excel COM (Windows only)."""
    import win32com.client
    import pythoncom
    pythoncom.CoInitialize()  # Required when called from a Flask worker thread
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            wb = excel.Workbooks.Open(os.path.abspath(input_path))
            wb.ExportAsFixedFormat(0, os.path.abspath(output_path))  # 0 = xlTypePDF
            wb.Close(False)
        finally:
            excel.Quit()
    finally:
        pythoncom.CoUninitialize()


def convert_with_win32com_ppt(input_path, output_path):
    """Convert PowerPoint to PDF using Microsoft PowerPoint COM (Windows only)."""
    import win32com.client
    import pythoncom
    pythoncom.CoInitialize()  # Required when called from a Flask worker thread
    try:
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        ppt_app.Visible = 1  # PowerPoint requires Visible=True to open files
        try:
            presentation = ppt_app.Presentations.Open(
                os.path.abspath(input_path), WithWindow=True
            )
            presentation.SaveAs(os.path.abspath(output_path), FileFormat=32)  # 32 = ppSaveAsPDF
            presentation.Close()
        finally:
            ppt_app.Quit()
    finally:
        pythoncom.CoUninitialize()


def office_convert(input_path, output_path, app_type):
    """
    2-tier conversion:
      1. win32com (Windows + MS Office) — best fidelity
      2. LibreOffice headless — good fidelity
    """
    errors = []

    # Tier 1: win32com (Windows only)
    if platform.system() == "Windows":
        try:
            if app_type == "word":
                convert_with_win32com_word(input_path, output_path)
            elif app_type == "excel":
                convert_with_win32com_excel(input_path, output_path)
            elif app_type == "ppt":
                convert_with_win32com_ppt(input_path, output_path)
            return
        except Exception as e:
            errors.append(f"win32com: {e}")

    # Tier 2: LibreOffice
    try:
        out_dir = os.path.dirname(output_path)
        pdf_path = convert_with_libreoffice(input_path, out_dir)
        if pdf_path != output_path:
            os.replace(pdf_path, output_path)
        return
    except Exception as e:
        errors.append(f"LibreOffice: {e}")

    raise RuntimeError("All conversion methods failed: " + " | ".join(errors))


# ── Word to PDF ──────────────────────────────────────────────────────────────

@app.route('/convert/word-to-pdf', methods=['POST'])
def word_to_pdf():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({"error": "No file selected"}), 400
    if not file.filename.lower().endswith(('.doc', '.docx')):
        return jsonify({"error": "Invalid file type. Only DOC/DOCX allowed."}), 400

    input_path = os.path.join(UPLOAD_FOLDER, file.filename)
    output_path = os.path.join(UPLOAD_FOLDER, "word_converted.pdf")
    file.save(input_path)

    try:
        office_convert(input_path, output_path, "word")
        return send_file(output_path, as_attachment=True, download_name="converted.pdf")
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        safe_remove(input_path)
        safe_remove(output_path)


# ── Excel to PDF ─────────────────────────────────────────────────────────────

@app.route('/convert/excel-to-pdf', methods=['POST'])
def excel_to_pdf():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({"error": "No file selected"}), 400
    if not file.filename.lower().endswith(('.xls', '.xlsx')):
        return jsonify({"error": "Invalid file type. Only XLS/XLSX allowed."}), 400

    input_path = os.path.join(UPLOAD_FOLDER, file.filename)
    output_path = os.path.join(UPLOAD_FOLDER, "excel_converted.pdf")
    file.save(input_path)

    try:
        office_convert(input_path, output_path, "excel")
        return send_file(output_path, as_attachment=True, download_name="converted.pdf")
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        safe_remove(input_path)
        safe_remove(output_path)


# ── PowerPoint to PDF ────────────────────────────────────────────────────────

@app.route('/convert/ppt-to-pdf', methods=['POST'])
def ppt_to_pdf():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({"error": "No file selected"}), 400
    if not file.filename.lower().endswith(('.ppt', '.pptx')):
        return jsonify({"error": "Invalid file type. Only PPT/PPTX allowed."}), 400

    input_path = os.path.join(UPLOAD_FOLDER, file.filename)
    output_path = os.path.join(UPLOAD_FOLDER, "ppt_converted.pdf")
    file.save(input_path)

    try:
        office_convert(input_path, output_path, "ppt")
        return send_file(output_path, as_attachment=True, download_name="converted.pdf")
    except Exception as e:
        return jsonify({"error": f"Conversion failed: {str(e)}"}), 500
    finally:
        safe_remove(input_path)
        safe_remove(output_path)


@app.route('/shutdown', methods=['GET', 'POST'])
def shutdown():
    """Gracefully shut down the Flask server."""
    def _stop():
        time.sleep(0.5)
        os._exit(0)
    threading.Thread(target=_stop, daemon=True).start()
    return jsonify({"message": "Server shutting down..."}), 200


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001, debug=False, use_reloader=False)
