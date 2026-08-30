import os
import uuid
import shutil
import subprocess

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
from pdf2docx import Converter


# ============================================================
# APP SETUP
# ============================================================

app = Flask(__name__)

CORS(app)

BASE_FOLDER = "/tmp/file_converter"

os.makedirs(
    BASE_FOLDER,
    exist_ok=True
)


# ============================================================
# REQUEST LOGGING
# ============================================================

@app.before_request
def log_request():

    print(
        f"REQUEST: {request.method} {request.path}",
        flush=True
    )


# ============================================================
# HOME
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
# HEALTH
# ============================================================

@app.route("/health", methods=["GET"])
def health():

    try:

        result = subprocess.run(
            ["libreoffice", "--version"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=20
        )

        version = (
            result.stdout.decode(
                "utf-8",
                errors="ignore"
            ).strip()
            or
            result.stderr.decode(
                "utf-8",
                errors="ignore"
            ).strip()
        )

        return jsonify({
            "status": "ok",
            "libreoffice": {
                "installed": result.returncode == 0,
                "version": version
            }
        })

    except Exception as error:

        return jsonify({
            "status": "error",
            "libreoffice": {
                "installed": False,
                "version": str(error)
            }
        }), 500


# ============================================================
# LIBREOFFICE HELPER
# ============================================================

def run_libreoffice(
    command,
    timeout=180
):

    print(
        "RUNNING:",
        " ".join(command),
        flush=True
    )

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

    print(
        "RETURN CODE:",
        result.returncode,
        flush=True
    )

    print(
        "STDOUT:",
        stdout,
        flush=True
    )

    print(
        "STDERR:",
        stderr,
        flush=True
    )

    return result, stdout, stderr


# ============================================================
# DOCX → PDF
# ============================================================

@app.route("/convert", methods=["POST"])
def docx_to_pdf():

    print(
        "DOCX TO PDF ROUTE STARTED",
        flush=True
    )

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

    safe_name = secure_filename(
        original_name
    )

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
        "input.docx"
    )

    uploaded_file.save(
        input_file
    )

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

        result, stdout, stderr = run_libreoffice(
            command
        )

        output_file = os.path.join(
            job_folder,
            "input.pdf"
        )

        if (
            result.returncode != 0
            or
            not os.path.exists(output_file)
        ):

            details = (
                stderr
                or
                stdout
                or
                "LibreOffice did not create the PDF."
            )

            return jsonify({

                "error":
                "DOCX to PDF conversion failed.",

                "details":
                details

            }), 500

        download_name = (
            os.path.splitext(
                safe_name
            )[0]
            +
            ".pdf"
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
            "DOCX to PDF conversion timed out."

        }), 500

    except Exception as error:

        return jsonify({

            "error":
            "DOCX to PDF conversion failed.",

            "details":
            str(error)

        }), 500

    finally:

        shutil.rmtree(
            job_folder,
            ignore_errors=True
        )


```python
# ============================================================
# PDF → WORD / DOCX
# Using pdf2docx instead of LibreOffice
# ============================================================

@app.route("/pdf-to-word", methods=["POST"])
def pdf_to_docx():

    print("==========================================", flush=True)
    print("PDF TO WORD ROUTE STARTED", flush=True)
    print("==========================================", flush=True)

    if "file" not in request.files:

        print("NO FILE RECEIVED", flush=True)

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


    os.makedirs(
        job_folder,
        exist_ok=True
    )


    input_file = os.path.join(
        job_folder,
        "input.pdf"
    )


    output_file = os.path.join(
        job_folder,
        "converted.docx"
    )


    uploaded_file.save(input_file)


    print(
        "PDF SAVED:",
        input_file,
        flush=True
    )


    print(
        "PDF SIZE:",
        os.path.getsize(input_file),
        "bytes",
        flush=True
    )


    try:

        # ----------------------------------------------------
        # Import pdf2docx
        # ----------------------------------------------------

        from pdf2docx import Converter


        print(
            "pdf2docx imported successfully",
            flush=True
        )


        # ----------------------------------------------------
        # Convert PDF → DOCX
        # ----------------------------------------------------

        print(
            "STARTING PDF → DOCX CONVERSION",
            flush=True
        )


        converter = Converter(
            input_file
        )


        try:

            converter.convert(
                output_file,
                start=0,
                end=None
            )

        finally:

            converter.close()


        # ----------------------------------------------------
        # Check output
        # ----------------------------------------------------

        print(
            "EXPECTED OUTPUT:",
            output_file,
            flush=True
        )


        print(
            "OUTPUT EXISTS:",
            os.path.exists(output_file),
            flush=True
        )


        if os.path.exists(output_file):

            output_size = os.path.getsize(
                output_file
            )

            print(
                "OUTPUT SIZE:",
                output_size,
                "bytes",
                flush=True
            )

        else:

            output_size = 0


        # ----------------------------------------------------
        # SUCCESS
        # ----------------------------------------------------

        if (
            os.path.exists(output_file)
            and
            output_size > 0
        ):

            download_name = (
                os.path.splitext(
                    original_name
                )[0]
                +
                ".docx"
            )


            print(
                "PDF → DOCX SUCCESS",
                flush=True
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


        # ----------------------------------------------------
        # FAILURE
        # ----------------------------------------------------

        print(
            "PDF → DOCX FAILED: Output file was not created.",
            flush=True
        )


        return jsonify({

            "error":
            "PDF to Word conversion failed.",

            "details":
            "The converter did not create the DOCX file."

        }), 500


    except Exception as error:

        print(
            "PDF → DOCX ERROR:",
            repr(error),
            flush=True
        )


        return jsonify({

            "error":
            "PDF to Word conversion failed.",

            "details":
            str(error)

        }), 500


    finally:

        print(
            "CLEANING JOB FOLDER:",
            job_folder,
            flush=True
        )


        shutil.rmtree(

            job_folder,

            ignore_errors=True

        )

