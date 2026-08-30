
import os
import uuid
import shutil
import subprocess
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename

# ============================================================
# APP SETUP
# ============================================================

app = Flask(__name__)
CORS(app)

@app.before_request
def log_request():
    print(
        f"REQUEST: {request.method} {request.path}",
        flush=True
    )
BASE_FOLDER = "/tmp/file_converter"
os.makedirs(BASE_FOLDER, exist_ok=True)


# ============================================================
# HOME / SERVER TEST
# ============================================================

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


# ============================================================
# LIBREOFFICE TEST
# ============================================================

def check_libreoffice():
    try:
        result = subprocess.run(
            ["libreoffice", "--version"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=20
        )

        return {
            "installed": result.returncode == 0,
            "version": (
                result.stdout.decode("utf-8", errors="ignore").strip()
                or result.stderr.decode("utf-8", errors="ignore").strip()
            )
        }

    except Exception as error:
        return {
            "installed": False,
            "version": str(error)
        }


@app.route("/health", methods=["GET"])
def health():
    lo = check_libreoffice()

    return jsonify({
        "status": "ok",
        "libreoffice": lo
    })


# ============================================================
# LIBREOFFICE HELPER
# ============================================================

def run_libreoffice(command, timeout=180):

    result = subprocess.run(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        timeout=timeout
    )

    stdout = result.stdout.decode(
        "utf-8",
        errors="ignore"
    ).strip()

    stderr = result.stderr.decode(
        "utf-8",
        errors="ignore"
    ).strip()

    return result, stdout, stderr


# ============================================================
# DOCX → PDF
# ============================================================

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

    safe_name = secure_filename(original_name)

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

        lo_profile = os.path.join(
            job_folder,
            "lo-profile"
        )

        os.makedirs(
            lo_profile,
            exist_ok=True
        )

        command = [
            "libreoffice",
            "--headless",
            "--nologo",
            "--nodefault",
            "--nofirststartwizard",
            "--norestore",
            f"-env:UserInstallation=file://{lo_profile}",
            "--convert-to",
            "pdf",
            "--outdir",
            job_folder,
            input_file
        ]

        result, stdout, stderr = run_libreoffice(command)

        output_file = os.path.join(
            job_folder,
            "input.pdf"
        )

        if result.returncode != 0 or not os.path.exists(output_file):

            details = (
                stderr
                or stdout
                or f"LibreOffice exited with code {result.returncode}"
            )

            return jsonify({
                "error": "DOCX to PDF conversion failed.",
                "details": details
            }), 500

        download_name = (
            os.path.splitext(safe_name)[0]
            + ".pdf"
        )

        return send_file(
            output_file,
            as_attachment=True,
            download_name=download_name,
            mimetype="application/pdf"
        )

    except subprocess.TimeoutExpired:

        return jsonify({
            "error":
            "Conversion timed out. The document may be too large or complex."
        }), 500

    except Exception as error:

        return jsonify({
            "error": "DOCX to PDF conversion failed.",
            "details": str(error)
        }), 500

    finally:

        shutil.rmtree(
            job_folder,
            ignore_errors=True
        )


# ============================================================
# PDF → WORD / DOCX
# ============================================================

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

    safe_name = secure_filename(original_name)

    job_id = str(uuid.uuid4())

    job_folder = os.path.join(
        BASE_FOLDER,
        job_id
    )

    os.makedirs(
        job_folder,
        exist_ok=True
    )

    input_file = os.path.join(
        job_folder,
        "input.pdf"
    )

    uploaded_file.save(input_file)

    try:

        # ----------------------------------------------------
        # Create unique LibreOffice profile
        # ----------------------------------------------------

        lo_profile = os.path.join(
            job_folder,
            "lo-profile"
        )

        os.makedirs(
            lo_profile,
            exist_ok=True
        )

        # ----------------------------------------------------
        # PDF → DOCX
        #
        # LibreOffice uses Draw to import PDF.
        # We explicitly specify the DOCX filter.
        # ----------------------------------------------------

        command = [
            "libreoffice",
            "--headless",
            "--nologo",
            "--nodefault",
            "--nofirststartwizard",
            "--norestore",
            f"-env:UserInstallation=file://{lo_profile}",
            "--convert-to",
            "docx:Office Open XML Text",
            "--outdir",
            job_folder,
            input_file
        ]

        result, stdout, stderr = run_libreoffice(
            command,
            timeout=180
        )

        output_file = os.path.join(
            job_folder,
            "input.docx"
        )

        # ----------------------------------------------------
        # Check return code AND output file
        # ----------------------------------------------------

        if result.returncode != 0 or not os.path.exists(output_file):

            details = (
                stderr
                or stdout
                or f"LibreOffice exited with code {result.returncode}"
            )

            return jsonify({
                "error": "PDF to Word conversion failed.",
                "details": details
            }), 500

        # ----------------------------------------------------
        # Return DOCX
        # ----------------------------------------------------

        download_name = (
            os.path.splitext(safe_name)[0]
            + ".docx"
        )

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
            "error":
            "PDF to Word conversion timed out. "
            "The PDF may be too large or complex."
        }), 500

    except Exception as error:

        return jsonify({
            "error": "PDF to Word conversion failed.",
            "details": str(error)
        }), 500

    finally:

        shutil.rmtree(
            job_folder,
            ignore_errors=True
        )


# ============================================================
# START SERVER
# ============================================================

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
