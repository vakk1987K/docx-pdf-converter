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


# =========================================================
# APPLICATION CONFIGURATION
# =========================================================

app = Flask(__name__)

# Allow Blogger / browser frontend requests
CORS(app)

# Maximum request size: 25 MB
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
            "PNG to PDF",
            "Multiple Images to PDF"
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

        version = (
            result.stdout.strip()
            if result.stdout
            else result.stderr.strip()
        )

        return jsonify({
            "success": True,
            "status": "ok",
            "libreoffice": {
                "installed": result.returncode == 0,
                "version": version
            }
        })

    except Exception as e:

        return jsonify({
            "success": False,
            "status": "error",
            "libreoffice": {
                "installed": False
            },
            "error": str(e)
        }), 500


# =========================================================
# LIBREOFFICE HELPER
# DOCX → PDF
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
            result.stderr or
            "LibreOffice conversion failed."
        )


# =========================================================
# DOCX → PDF
# Endpoint: POST /convert
# =========================================================

@app.route("/convert", methods=["POST"])
def convert_word_to_pdf():

    job_id = uuid.uuid4().hex

    job_folder = os.path.join(
        BASE_FOLDER,
        job_id
    )

    os.makedirs(
        job_folder,
        exist_ok=True
    )

    try:

        # -------------------------------------------------
        # Check upload
        # -------------------------------------------------

        if "file" not in request.files:

            return jsonify({
                "success": False,
                "error": "Please upload a DOCX file."
            }), 400

        uploaded_file = request.files["file"]

        if (
            not uploaded_file
            or uploaded_file.filename == ""
        ):

            return jsonify({
                "success": False,
                "error": "No file selected."
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

        if extension != ".docx":

            return jsonify({
                "success": False,
                "error": "Only DOCX files are supported."
            }), 400

        # -------------------------------------------------
        # Save input
        # -------------------------------------------------

        input_path = os.path.join(
            job_folder,
            "input.docx"
        )

        uploaded_file.save(input_path)

        # -------------------------------------------------
        # Convert
        # -------------------------------------------------

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
                "PDF conversion failed. "
                "Output file was not created."
            )

        if os.path.getsize(output_path) == 0:

            raise RuntimeError(
                "Generated PDF is empty."
            )

        # -------------------------------------------------
        # Send PDF
        # -------------------------------------------------

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
            "DOCX to PDF conversion error:",
            str(e)
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
# Endpoint: POST /pdf-to-word
# =========================================================

@app.route("/pdf-to-word", methods=["POST"])
def convert_pdf_to_word():

    job_id = uuid.uuid4().hex

    job_folder = os.path.join(
        BASE_FOLDER,
        job_id
    )

    os.makedirs(
        job_folder,
        exist_ok=True
    )

    try:

        # -------------------------------------------------
        # Check upload
        # -------------------------------------------------

        if "file" not in request.files:

            return jsonify({
                "success": False,
                "error": "Please upload a PDF file."
            }), 400

        uploaded_file = request.files["file"]

        if (
            not uploaded_file
            or uploaded_file.filename == ""
        ):

            return jsonify({
                "success": False,
                "error": "No PDF file selected."
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

        if extension != ".pdf":

            return jsonify({
                "success": False,
                "error": "Only PDF files are supported."
            }), 400

        # -------------------------------------------------
        # Save PDF
        # -------------------------------------------------

        input_path = os.path.join(
            job_folder,
            "input.pdf"
        )

        output_path = os.path.join(
            job_folder,
            "converted.docx"
        )

        uploaded_file.save(input_path)

        # -------------------------------------------------
        # Verify PDF before conversion
        # -------------------------------------------------

        try:

            reader = PdfReader(input_path)

            if len(reader.pages) == 0:

                raise ValueError(
                    "PDF contains no pages."
                )

        except Exception as pdf_error:

            print(
                "PDF validation error:",
                str(pdf_error)
            )

            return jsonify({
                "success": False,
                "error": "The uploaded file is not a valid PDF."
            }), 400

        # -------------------------------------------------
        # PDF → DOCX
        # -------------------------------------------------

        converter = Converter(input_path)

        try:

            converter.convert(
                output_path
            )

        finally:

            converter.close()

        # -------------------------------------------------
        # Verify output
        # -------------------------------------------------

        if not os.path.exists(output_path):

            raise RuntimeError(
                "DOCX file was not created."
            )

        if os.path.getsize(output_path) == 0:

            raise RuntimeError(
                "Generated DOCX file is empty."
            )

        # -------------------------------------------------
        # Send DOCX
        # -------------------------------------------------

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
            "PDF to Word conversion error:",
            str(e)
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
# Endpoint: POST /pdf-to-text
# =========================================================

@app.route("/pdf-to-text", methods=["POST"])
def convert_pdf_to_text():

    job_id = uuid.uuid4().hex

    job_folder = os.path.join(
        BASE_FOLDER,
        job_id
    )

    os.makedirs(
        job_folder,
        exist_ok=True
    )

    try:

        # -------------------------------------------------
        # Check upload
        # -------------------------------------------------

        if "file" not in request.files:

            return jsonify({
                "success": False,
                "error": "Please upload a PDF file."
            }), 400

        uploaded_file = request.files["file"]

        if (
            not uploaded_file
            or uploaded_file.filename == ""
        ):

            return jsonify({
                "success": False,
                "error": "No PDF file selected."
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

        if extension != ".pdf":

            return jsonify({
                "success": False,
                "error": "Only PDF files are supported."
            }), 400

        # -------------------------------------------------
        # Save PDF
        # -------------------------------------------------

        input_path = os.path.join(
            job_folder,
            "input.pdf"
        )

        output_path = os.path.join(
            job_folder,
            "converted.txt"
        )

        uploaded_file.save(input_path)

        # -------------------------------------------------
        # Read PDF
        # -------------------------------------------------

        reader = PdfReader(
            input_path
        )

        if len(reader.pages) == 0:

            return jsonify({
                "success": False,
                "error": "The PDF contains no pages."
            }), 400

        extracted_text = []

        for page_number, page in enumerate(
            reader.pages,
            start=1
        ):

            try:

                text = (
                    page.extract_text()
                    or ""
                )

            except Exception as page_error:

                print(
                    f"Page {page_number} "
                    f"text extraction error:",
                    str(page_error)
                )

                text = ""

            extracted_text.append(
                f"--- Page {page_number} ---\n"
                f"{text.strip()}"
            )

        final_text = (
            "\n\n".join(
                extracted_text
            ).strip()
        )

        # -------------------------------------------------
        # No readable text
        # -------------------------------------------------

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

        # -------------------------------------------------
        # Save TXT
        # -------------------------------------------------

        with open(
            output_path,
            "w",
            encoding="utf-8"
        ) as text_file:

            text_file.write(
                final_text
            )

        # -------------------------------------------------
        # Send TXT
        # -------------------------------------------------

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
            "PDF to Text conversion error:",
            str(e)
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
# Endpoint: POST /image-to-pdf
# Supports JPG / JPEG / PNG
# =========================================================

@app.route("/image-to-pdf", methods=["POST"])
def convert_image_to_pdf():

    job_id = uuid.uuid4().hex

    job_folder = os.path.join(
        BASE_FOLDER,
        job_id
    )

    os.makedirs(
        job_folder,
        exist_ok=True
    )

    try:

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
        # Save image
        # -------------------------------------------------

        input_path = os.path.join(
            job_folder,
            "input" + extension
        )

        output_path = os.path.join(
            job_folder,
            "converted-image.pdf"
        )

        uploaded_file.save(
            input_path
        )

        # -------------------------------------------------
        # Validate image
        # -------------------------------------------------

        try:

            with Image.open(
                input_path
            ) as test_image:

                test_image.verify()

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
        # Open image again
        # -------------------------------------------------

        with Image.open(
            input_path
        ) as image:

            # Correct phone-camera orientation
            image = ImageOps.exif_transpose(
                image
            )

            # -------------------------------------------------
            # Convert transparent images to white background
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

                alpha = image.getchannel(
                    "A"
                )

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

                alpha = image.getchannel(
                    "A"
                )

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
            # Create PDF
            # -------------------------------------------------

            image.save(
                output_path,
                "PDF",
                resolution=100.0
            )

        # -------------------------------------------------
        # Verify PDF
        # -------------------------------------------------

        if not os.path.exists(
            output_path
        ):

            raise RuntimeError(
                "PDF file was not created."
            )

        if os.path.getsize(
            output_path
        ) == 0:

            raise RuntimeError(
                "Generated PDF is empty."
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

        @response.call_on_close
        def cleanup():

            shutil.rmtree(
                job_folder,
                ignore_errors=True
            )

        return response

    except Exception as e:

        print(
            "Image to PDF conversion error:",
            str(e)
        )

        shutil.rmtree(
            job_folder,
            ignore_errors=True
        )

        return jsonify({
            "success": False,
            "error": "Unable to convert the image to PDF."
        }), 500


# =========================================================
# MULTIPLE IMAGES → PDF
# Endpoint: POST /images-to-pdf
#
# Maximum:
#   20 images
#   25 MB total
#
# Supported:
#   JPG
#   JPEG
#   PNG
# =========================================================

@app.route("/images-to-pdf", methods=["POST"])
def convert_multiple_images_to_pdf():

    MAX_IMAGES = 20
    MAX_TOTAL_SIZE = 25 * 1024 * 1024

    allowed_extensions = {
        ".jpg",
        ".jpeg",
        ".png"
    }

    job_id = uuid.uuid4().hex

    job_folder = os.path.join(
        BASE_FOLDER,
        f"multi-{job_id}"
    )

    os.makedirs(
        job_folder,
        exist_ok=True
    )

    try:

        # -------------------------------------------------
        # Get multiple files
        # -------------------------------------------------

        files = request.files.getlist(
            "files"
        )

        if not files:

            # Also accept "file" if frontend
            # sends multiple file inputs using file.
            files = request.files.getlist(
                "file"
            )

        if not files:

            return jsonify({
                "success": False,
                "error": (
                    "Please select at least one "
                    "JPG, JPEG or PNG image."
                )
            }), 400

        # -------------------------------------------------
        # Maximum number of images
        # -------------------------------------------------

        if len(files) > MAX_IMAGES:

            return jsonify({
                "success": False,
                "error": (
                    f"Maximum {MAX_IMAGES} "
                    "images are allowed."
                )
            }), 400

        # -------------------------------------------------
        # Calculate total upload size
        # -------------------------------------------------

        total_size = 0

        for uploaded_file in files:

            if (
                not uploaded_file
                or not uploaded_file.filename
            ):

                return jsonify({
                    "success": False,
                    "error": (
                        "One of the selected "
                        "files is invalid."
                    )
                }), 400

            filename = secure_filename(
                uploaded_file.filename
            )

            if not filename:

                return jsonify({
                    "success": False,
                    "error": "Invalid filename detected."
                }), 400

            extension = os.path.splitext(
                filename
            )[1].lower()

            if extension not in allowed_extensions:

                return jsonify({
                    "success": False,
                    "error": (
                        "Only JPG, JPEG and PNG "
                        "images are supported."
                    )
                }), 400

            # Check actual uploaded stream size
            uploaded_file.stream.seek(
                0,
                os.SEEK_END
            )

            file_size = (
                uploaded_file.stream.tell()
            )

            uploaded_file.stream.seek(
                0
            )

            total_size += file_size

            if total_size > MAX_TOTAL_SIZE:

                return jsonify({
                    "success": False,
                    "error": (
                        "The total image size "
                        "must not exceed 25 MB."
                    )
                }), 413

        # -------------------------------------------------
        # Process every image
        # -------------------------------------------------

        converted_images = []

        for index, uploaded_file in enumerate(
            files
        ):

            input_path = os.path.join(
                job_folder,
                f"input_{index}.upload"
            )

            uploaded_file.save(
                input_path
            )

            # -------------------------------------------------
            # Verify image contents
            # -------------------------------------------------

            try:

                with Image.open(
                    input_path
                ) as test_image:

                    test_image.verify()

            except Exception:

                return jsonify({
                    "success": False,
                    "error": (
                        f"Image {index + 1} "
                        "is not a valid JPG, JPEG "
                        "or PNG file."
                    )
                }), 400

            # -------------------------------------------------
            # Process image
            # -------------------------------------------------

            try:

                with Image.open(
                    input_path
                ) as image:

                    # Correct camera orientation
                    image = ImageOps.exif_transpose(
                        image
                    )

                    # -------------------------------------------------
                    # Handle transparent / palette images
                    # -------------------------------------------------

                    if image.mode in (
                        "RGBA",
                        "LA",
                        "P"
                    ):

                        if image.mode == "P":

                            image = image.convert(
                                "RGBA"
                            )

                        background = Image.new(
                            "RGB",
                            image.size,
                            "white"
                        )

                        if "A" in image.getbands():

                            background.paste(
                                image,
                                mask=image.getchannel(
                                    "A"
                                )
                            )

                        else:

                            background.paste(
                                image
                            )

                        image = background

                    else:

                        image = image.convert(
                            "RGB"
                        )

                    # -------------------------------------------------
                    # Normalize to JPEG
                    # -------------------------------------------------

                    output_image = os.path.join(
                        job_folder,
                        f"page_{index:03d}.jpg"
                    )

                    image.save(
                        output_image,
                        "JPEG",
                        quality=92,
                        optimize=True
                    )

                    converted_images.append(
                        output_image
                    )

            except Exception as image_error:

                print(
                    "Image processing error:",
                    str(image_error)
                )

                return jsonify({
                    "success": False,
                    "error": (
                        f"Could not process "
                        f"image {index + 1}."
                    )
                }), 400

        # -------------------------------------------------
        # Make sure images exist
        # -------------------------------------------------

        if not converted_images:

            return jsonify({
                "success": False,
                "error": "No valid images were found."
            }), 400

        # -------------------------------------------------
        # Create ONE PDF
        # in exact upload order
        # -------------------------------------------------

        pdf_path = os.path.join(
            job_folder,
            "multiple-images.pdf"
        )

        pdf_images = []

        try:

            for image_path in converted_images:

                image = Image.open(
                    image_path
                )

                image.load()

                pdf_images.append(
                    image
                )

            first_image = pdf_images[0]

            remaining_images = (
                pdf_images[1:]
            )

            first_image.save(
                pdf_path,
                "PDF",
                resolution=100.0,
                save_all=True,
                append_images=remaining_images
            )

        finally:

            for image in pdf_images:

                try:
                    image.close()

                except Exception:
                    pass

        # -------------------------------------------------
        # Verify PDF
        # -------------------------------------------------

        if not os.path.exists(
            pdf_path
        ):

            raise RuntimeError(
                "PDF file was not created."
            )

        if os.path.getsize(
            pdf_path
        ) == 0:

            raise RuntimeError(
                "Generated PDF is empty."
            )

        print(
            "Multiple images converted successfully:",
            len(converted_images),
            "images"
        )

        # -------------------------------------------------
        # Send PDF
        # -------------------------------------------------

        response = send_file(
            pdf_path,
            mimetype="application/pdf",
            as_attachment=True,
            download_name="multiple-images.pdf"
        )

        # -------------------------------------------------
        # Cleanup after response
        # -------------------------------------------------

        @response.call_on_close
        def cleanup():

            try:

                shutil.rmtree(
                    job_folder,
                    ignore_errors=True
                )

            except Exception as cleanup_error:

                print(
                    "Cleanup error:",
                    str(cleanup_error)
                )

        return response

    except Exception as e:

        print(
            "Multiple image conversion error:",
            str(e)
        )

        shutil.rmtree(
            job_folder,
            ignore_errors=True
        )

        return jsonify({
            "success": False,
            "error": (
                "Unable to create the PDF. "
                "Please try again."
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
