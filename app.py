import os
import uuid
import shutil
import subprocess

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename

from pdf2docx import Converter
from pypdf import PdfReader

from PIL import Image, ImageOps
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader


# =========================================================
# APP CONFIGURATION
# =========================================================

app = Flask(__name__)

# Allow Blogger / browser frontend requests
CORS(app)

# Maximum upload size: 25 MB
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

# Temporary working directory
BASE_FOLDER = "/tmp/file_converter"
os.makedirs(BASE_FOLDER, exist_ok=True)


# =========================================================
# SECURITY HEADERS
# =========================================================

@app.after_request
def add_security_headers(response):
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "SAMEORIGIN"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    response.headers["Cache-Control"] = "no-store"

    return response


# =========================================================
# REQUEST LOGGING
# =========================================================

@app.before_request
def log_request():
    print(
        f"{request.method} {request.path} "
        f"from {request.remote_addr}"
    )


# =========================================================
# FILE SIZE ERROR
# =========================================================

@app.errorhandler(413)
def request_entity_too_large(error):
    return jsonify({
        "success": False,
        "error": "File is too large. Maximum allowed size is 25 MB."
    }), 413


# =========================================================
# HOME
# =========================================================

@app.route("/", methods=["GET"])
def home():
    return jsonify({
        "name": "PDF Word Converter API",
        "status": "online",
        "features": [
            "DOCX to PDF",
            "PDF to DOCX",
            "PDF to TXT",
            "JPG to PDF",
            "JPEG to PDF",
            "PNG to PDF"
        ]
    })


# =========================================================
# HEALTH CHECK
# =========================================================

@app.route("/health", methods=["GET"])
def health():
    try:
        result = subprocess.run(
            ["libreoffice", "--version"],
            capture_output=True,
            text=True,
            timeout=10
        )

        version = result.stdout.strip()

        return jsonify({
            "status": "ok",
            "libreoffice": {
                "installed": result.returncode == 0,
                "version": version
            }
        })

    except Exception as e:
        return jsonify({
            "status": "error",
            "libreoffice": {
                "installed": False
            },
            "error": str(e)
        }), 500


# =========================================================
# LIBREOFFICE HELPER
# =========================================================

def run_libreoffice(input_file, output_folder):
    command = [
        "libreoffice",
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        output_folder,
        input_file
    ]

    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        timeout=180
    )

    print("LibreOffice stdout:", result.stdout)
    print("LibreOffice stderr:", result.stderr)

    if result.returncode != 0:
        raise RuntimeError(
            result.stderr or "LibreOffice conversion failed."
        )


# =========================================================
# DOCX → PDF
# EXISTING ROUTE
# =========================================================

@app.route("/convert", methods=["POST"])
def convert_word_to_pdf():

    job_id = str(uuid.uuid4())
    job_folder = os.path.join(BASE_FOLDER, job_id)

    try:
        os.makedirs(job_folder, exist_ok=True)

        if "file" not in request.files:
            return jsonify({
                "success": False,
                "error": "Please upload a DOCX file."
            }), 400

        uploaded_file = request.files["file"]

        if not uploaded_file or uploaded_file.filename == "":
            return jsonify({
                "success": False,
                "error": "No file selected."
            }), 400

        filename = secure_filename(uploaded_file.filename)

        if not filename:
            return jsonify({
                "success": False,
                "error": "Invalid filename."
            }), 400

        extension = os.path.splitext(filename)[1].lower()

        if extension != ".docx":
            return jsonify({
                "success": False,
                "error": "Only DOCX files are supported."
            }), 400

        input_path = os.path.join(job_folder, "input.docx")

        uploaded_file.save(input_path)

        run_libreoffice(
            input_path,
            job_folder
        )

        output_path = os.path.join(
            job_folder,
            "input.pdf"
        )

        if not os.path.exists(output_path):
            raise RuntimeError(
                "PDF conversion failed. Output file was not created."
            )

        response = send_file(
            output_path,
            mimetype="application/pdf",
            as_attachment=True,
            download_name="converted.pdf"
        )

        @response.call_on_close
        def cleanup():
            shutil.rmtree(
                job_folder,
                ignore_errors=True
            )

        return response

    except Exception as e:

        print(
            f"DOCX to PDF conversion error: {str(e)}"
        )

        shutil.rmtree(
            job_folder,
            ignore_errors=True
        )

        return jsonify({
            "success": False,
            "error": "Unable to convert DOCX to PDF."
        }), 500


# =========================================================
# PDF → WORD
# EXISTING ROUTE
# =========================================================

@app.route("/pdf-to-word", methods=["POST"])
def convert_pdf_to_word():

    job_id = str(uuid.uuid4())
    job_folder = os.path.join(BASE_FOLDER, job_id)

    try:
        os.makedirs(job_folder, exist_ok=True)

        if "file" not in request.files:
            return jsonify({
                "success": False,
                "error": "Please upload a PDF file."
            }), 400

        uploaded_file = request.files["file"]

        if not uploaded_file or uploaded_file.filename == "":
            return jsonify({
                "success": False,
                "error": "No PDF file selected."
            }), 400

        filename = secure_filename(
            uploaded_file.filename
        )

        if not filename:
            return jsonify({
                "success": False,
                "error": "Invalid filename."
            }), 400

        extension = os.path.splitext(filename)[1].lower()

        if extension != ".pdf":
            return jsonify({
                "success": False,
                "error": "Only PDF files are supported."
            }), 400

        input_path = os.path.join(
            job_folder,
            "input.pdf"
        )

        output_path = os.path.join(
            job_folder,
            "converted.docx"
        )

        uploaded_file.save(input_path)

        converter = Converter(input_path)

        try:
            converter.convert(output_path)

        finally:
            converter.close()

        if not os.path.exists(output_path):
            raise RuntimeError(
                "DOCX file was not created."
            )

        response = send_file(
            output_path,
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            as_attachment=True,
            download_name="converted.docx"
        )

        @response.call_on_close
        def cleanup():
            shutil.rmtree(
                job_folder,
                ignore_errors=True
            )

        return response

    except Exception as e:

        print(
            f"PDF to Word conversion error: {str(e)}"
        )

        shutil.rmtree(
            job_folder,
            ignore_errors=True
        )

        return jsonify({
            "success": False,
            "error": "Unable to convert PDF to Word."
        }), 500


# =========================================================
# PDF → TEXT
# EXISTING ROUTE
# =========================================================

@app.route("/pdf-to-text", methods=["POST"])
def convert_pdf_to_text():

    job_id = str(uuid.uuid4())
    job_folder = os.path.join(BASE_FOLDER, job_id)

    try:
        os.makedirs(job_folder, exist_ok=True)

        if "file" not in request.files:
            return jsonify({
                "success": False,
                "error": "Please upload a PDF file."
            }), 400

        uploaded_file = request.files["file"]

        if not uploaded_file or uploaded_file.filename == "":
            return jsonify({
                "success": False,
                "error": "No PDF file selected."
            }), 400

        filename = secure_filename(
            uploaded_file.filename
        )

        if not filename:
            return jsonify({
                "success": False,
                "error": "Invalid filename."
            }), 400

        extension = os.path.splitext(filename)[1].lower()

        if extension != ".pdf":
            return jsonify({
                "success": False,
                "error": "Only PDF files are supported."
            }), 400

        input_path = os.path.join(
            job_folder,
            "input.pdf"
        )

        output_path = os.path.join(
            job_folder,
            "converted.txt"
        )

        uploaded_file.save(input_path)

        reader = PdfReader(input_path)

        extracted_text = []

        for page_number, page in enumerate(
            reader.pages,
            start=1
        ):

            try:
                text = page.extract_text() or ""

            except Exception as page_error:

                print(
                    f"Page {page_number} text extraction error: "
                    f"{page_error}"
                )

                text = ""

            extracted_text.append(
                f"--- Page {page_number} ---\n{text.strip()}"
            )

        final_text = "\n\n".join(
            extracted_text
        ).strip()

        if not final_text:
            shutil.rmtree(
                job_folder,
                ignore_errors=True
            )

            return jsonify({
                "success": False,
                "error": (
                    "No readable text was found in this PDF. "
                    "The PDF may be scanned or image-based."
                )
            }), 422

        with open(
            output_path,
            "w",
            encoding="utf-8"
        ) as text_file:

            text_file.write(final_text)

        response = send_file(
            output_path,
            mimetype="text/plain; charset=utf-8",
            as_attachment=True,
            download_name="converted.txt"
        )

        @response.call_on_close
        def cleanup():
            shutil.rmtree(
                job_folder,
                ignore_errors=True
            )

        return response

    except Exception as e:

        print(
            f"PDF to Text conversion error: {str(e)}"
        )

        shutil.rmtree(
            job_folder,
            ignore_errors=True
        )

        return jsonify({
            "success": False,
            "error": "Unable to extract text from PDF."
        }), 500


# =========================================================
# IMAGE → PDF
# JPG / JPEG / PNG
# NEW ROUTE
# =========================================================

@app.route("/image-to-pdf", methods=["POST"])
def convert_image_to_pdf():

    job_id = str(uuid.uuid4())
    job_folder = os.path.join(
        BASE_FOLDER,
        job_id
    )

    try:
        os.makedirs(
            job_folder,
            exist_ok=True
        )

        # -------------------------------------------------
        # Check upload
        # -------------------------------------------------

        if "file" not in request.files:
            return jsonify({
                "success": False,
                "error": "Please upload an image file."
            }), 400

        uploaded_file = request.files["file"]

        if (
            not uploaded_file
            or uploaded_file.filename == ""
        ):
            return jsonify({
                "success": False,
                "error": "No image file selected."
            }), 400

        # -------------------------------------------------
        # Secure filename
        # -------------------------------------------------

        filename = secure_filename(
            uploaded_file.filename
        )

        if not filename:
            return jsonify({
                "success": False,
                "error": "Invalid filename."
            }), 400

        extension = os.path.splitext(
            filename
        )[1].lower()

        # -------------------------------------------------
        # Allowed formats
        # -------------------------------------------------

        allowed_extensions = {
            ".jpg",
            ".jpeg",
            ".png"
        }

        if extension not in allowed_extensions:
            return jsonify({
                "success": False,
                "error": (
                    "Only JPG, JPEG and PNG "
                    "images are supported."
                )
            }), 400

        # -------------------------------------------------
        # Save uploaded image
        # -------------------------------------------------

        input_path = os.path.join(
            job_folder,
            "input" + extension
        )

        output_path = os.path.join(
            job_folder,
            "converted-image.pdf"
        )

        uploaded_file.save(input_path)

        # -------------------------------------------------
        # Validate image
        # -------------------------------------------------

        try:

            with Image.open(input_path) as image:

                image.verify()

        except Exception:

            shutil.rmtree(
                job_folder,
                ignore_errors=True
            )

            return jsonify({
                "success": False,
                "error": (
                    "The uploaded file is not "
                    "a valid JPG, JPEG or PNG image."
                )
            }), 400

        # -------------------------------------------------
        # Open image again after verify()
        # -------------------------------------------------

        with Image.open(input_path) as image:

            # Correct phone-camera orientation
            image = ImageOps.exif_transpose(
                image
            )

            # -------------------------------------------------
            # Convert to RGB
            #
            # PDF doesn't directly support all PNG modes.
            # Transparent PNG areas become white.
            # -------------------------------------------------

            if image.mode in (
                "RGBA",
                "LA"
            ):

                background = Image.new(
                    "RGB",
                    image.size,
                    "white"
                )

                alpha = image.getchannel("A")

                background.paste(
                    image,
                    mask=alpha
                )

                image = background

            elif image.mode == "P":

                image = image.convert(
                    "RGBA"
                )

                background = Image.new(
                    "RGB",
                    image.size,
                    "white"
                )

                alpha = image.getchannel("A")

                background.paste(
                    image,
                    mask=alpha
                )

                image = background

            else:

                image = image.convert(
                    "RGB"
                )

            # -------------------------------------------------
            # Save as PDF
            # -------------------------------------------------

            image.save(
                output_path,
                "PDF",
                resolution=100.0
            )

        # -------------------------------------------------
        # Verify PDF exists
        # -------------------------------------------------

        if not os.path.exists(
            output_path
        ):

            raise RuntimeError(
                "PDF file was not created."
            )

        # -------------------------------------------------
        # Send PDF
        # -------------------------------------------------

        response = send_file(
            output_path,
            mimetype="application/pdf",
            as_attachment=True,
            download_name="converted-image.pdf"
        )

        # -------------------------------------------------
        # Cleanup after response
        # -------------------------------------------------

        @response.call_on_close
        def cleanup():

            shutil.rmtree(
                job_folder,
                ignore_errors=True
            )

        return response

    except Exception as e:

        print(
            f"Image to PDF conversion error: {str(e)}"
        )

        shutil.rmtree(
            job_folder,
            ignore_errors=True
        )

        return jsonify({
            "success": False,
            "error": (
                "Unable to convert the image to PDF."
            )
        }), 500


# =========================================================
# 404 HANDLER
# =========================================================

@app.errorhandler(404)
def not_found(error):
    return jsonify({
        "success": False,
        "error": "Endpoint not found."
    }), 404


# =========================================================
# 500 HANDLER
# =========================================================

@app.errorhandler(500)
def internal_server_error(error):
    return jsonify({
        "success": False,
        "error": "An internal server error occurred."
    }), 500


# =========================================================
# LOCAL DEVELOPMENT
# Render uses Gunicorn
# =========================================================

if __name__ == "__main__":

    port = int(
        os.environ.get(
            "PORT",
            10000
        )
    )

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False
    )
