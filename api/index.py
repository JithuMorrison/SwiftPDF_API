from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import time
import nbformat
import tempfile
from fpdf import FPDF

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


# Convert IPython Notebook to PDF
@app.route('/convert/ipynb-to-pdf', methods=['POST'])
def ipynb_to_pdf():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No file selected"}), 400

    input_path = os.path.join(UPLOAD_FOLDER, file.filename)
    output_path = os.path.join(UPLOAD_FOLDER, "converted.pdf")
    file.save(input_path)

    try:
        with open(input_path, 'r', encoding='utf-8') as f:
            notebook = nbformat.read(f, as_version=4)

        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        for cell in notebook.cells:
            if cell.cell_type == "markdown":
                pdf.set_font("Arial", style='B', size=14)
                pdf.multi_cell(0, 10, cell.source)
                pdf.ln(5)
            elif cell.cell_type == "code":
                pdf.set_font("Courier", size=10)
                pdf.multi_cell(0, 8, cell.source)
                pdf.ln(5)

        pdf.output(output_path)
        return send_file(output_path, as_attachment=True, download_name="converted.pdf")

    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        safe_remove(input_path)
        safe_remove(output_path)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
