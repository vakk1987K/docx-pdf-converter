```python
import os
import uuid
import shutil
import subprocess

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS


app = Flask(__name__)

# Allow Blogger to communicate with this server
CORS(app)


BASE_FOLDER = "/tmp/file_converter"

os.makedirs(BASE_FOLDER, exist_ok=True)


# ---------------------------------------------------------
# HOME / SERVER TEST
# ---------------------------------------------------------

@app.route("/", methods=["GET"])
def home():

    return jsonify({
        "status": "online",
        "service": "TeluguTech777 File Converter",
        "features": [
            "DOCX to PDF",
            "PDF to DOCX"
        ]
    })


# ---------------------------------------------------------
# DOCX → PDF
# ---------------------------------------------------------

@app.route("/convert", methods=["POST"])
def docx_to_pdf():

    if "file" not in request.files:

        return jsonify({
            "error": "No file uploaded."
        }), 400


    uploaded_file = request.files["file"]


    if uploaded_file.filename == "":

        return jsonify({
            "error": "No file selected."
        }), 400


    original_name = uploaded_file.filename


    if not original_name.lower().endswith(".docx"):

        return jsonify({
            "error": "Only DOCX files are supported."
        }), 400


    job_id = str(uuid.uuid4())

    job_folder = os.path.join(
        BASE_FOLDER,
        job_id
    )

    os.makedirs(job_folder, exist_ok=True)


    input_file = os.path.join(
        job_folder,
        "input.docx"
    )


    uploaded_file.save(input_file)


    try:

        command = [
            "libreoffice",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            job_folder,
            input_file
        ]


        result = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=180
        )


        output_file = os.path.join(
            job_folder,
            "input.pdf"
        )


        if not os.path.exists(output_file):

            error_message = (
                result.stderr.decode(
                    "utf-8",
                    errors="ignore"
                )
                or
                result.stdout.decode(
                    "utf-8",
                    errors="ignore"
                )
            )


            return jsonify({
                "error": "DOCX to PDF conversion failed.",
                "details": error_message
            }), 500


        download_name = os.path.splitext(
            original_name
        )[0] + ".pdf"


        return send_file(
            output_file,
            as_attachment=True,
            download_name=download_name,
            mimetype="application/pdf"
        )


    except subprocess.TimeoutExpired:

        return jsonify({
            "error": "Conversion timed out. The document may be too large or complex."
        }), 500


    except Exception as error:

        return jsonify({
            "error": str(error)
        }), 500


    finally:

        # Delete temporary files
        shutil.rmtree(
            job_folder,
            ignore_errors=True
        )


# ---------------------------------------------------------
# PDF → WORD / DOCX
# ---------------------------------------------------------

@app.route("/pdf-to-word", methods=["POST"])
def pdf_to_docx():

    if "file" not in request.files:

        return jsonify({
            "error": "No PDF file uploaded."
        }), 400


    uploaded_file = request.files["file"]


    if uploaded_file.filename == "":

        return jsonify({
            "error": "No file selected."
        }), 400


    original_name = uploaded_file.filename


    if not original_name.lower().endswith(".pdf"):

        return jsonify({
            "error": "Only PDF files are supported."
        }), 400


    job_id = str(uuid.uuid4())

    job_folder = os.path.join(
        BASE_FOLDER,
        job_id
    )

    os.makedirs(job_folder, exist_ok=True)


    input_file = os.path.join(
        job_folder,
        "input.pdf"
    )


    uploaded_file.save(input_file)


    try:

        # LibreOffice opens PDF as a Draw document
        # and exports it as DOCX.
        command = [
            "libreoffice",
            "--headless",
            "--convert-to",
            "docx",
            "--outdir",
            job_folder,
            input_file
        ]


        result = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=180
        )


        output_file = os.path.join(
            job_folder,
            "input.docx"
        )


        if not os.path.exists(output_file):

            error_message = (
                result.stderr.decode(
                    "utf-8",
                    errors="ignore"
                )
                or
                result.stdout.decode(
                    "utf-8",
                    errors="ignore"
                )
            )


            return jsonify({
                "error": "PDF to Word conversion failed.",
                "details": error_message
            }), 500


        download_name = os.path.splitext(
            original_name
        )[0] + ".docx"


        return send_file(
            output_file,
            as_attachment=True,
            download_name=download_name,
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            )
        )


    except subprocess.TimeoutExpired:

        return jsonify({
            "error": "Conversion timed out."
        }), 500


    except Exception as error:

        return jsonify({
            "error": str(error)
        }), 500


    finally:

        # Delete temporary files
        shutil.rmtree(
            job_folder,
            ignore_errors=True
        )


# ---------------------------------------------------------
# START SERVER
# ---------------------------------------------------------

if __name__ == "__main__":

    port = int(
        os.environ.get(
            "PORT",
            10000
        )
    )


    app.run(
        host="0.0.0.0",
        port=port
    )
```
