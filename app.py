import os
import uuid
import subprocess
import shutil

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS

app = Flask(__name__)

CORS(app)

UPLOAD_FOLDER = "/tmp/docx_converter"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


@app.route("/", methods=["GET"])
def home():

    return jsonify({
        "status": "online",
        "service": "DOCX to PDF Converter"
    })


@app.route("/convert", methods=["POST"])
def convert():

    if "file" not in request.files:

        return jsonify({
            "error": "No file uploaded"
        }), 400


    uploaded_file = request.files["file"]


    if uploaded_file.filename == "":

        return jsonify({
            "error": "No file selected"
        }), 400


    filename = uploaded_file.filename


    if not filename.lower().endswith(".docx"):

        return jsonify({
            "error": "Only DOCX files are supported"
        }), 400


    job_id = str(uuid.uuid4())

    job_folder = os.path.join(
        UPLOAD_FOLDER,
        job_id
    )

    os.makedirs(job_folder, exist_ok=True)


    input_file = os.path.join(
        job_folder,
        filename
    )


    uploaded_file.save(input_file)


    try:

        result = subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                job_folder,
                input_file
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=120
        )


        pdf_filename = os.path.splitext(
            filename
        )[0] + ".pdf"


        pdf_file = os.path.join(
            job_folder,
            pdf_filename
        )


        if not os.path.exists(pdf_file):

            return jsonify({
                "error": "PDF conversion failed",
                "details": result.stderr.decode(
                    "utf-8",
                    errors="ignore"
                )
            }), 500


        return send_file(
            pdf_file,
            as_attachment=True,
            download_name=pdf_filename,
            mimetype="application/pdf"
        )


    except subprocess.TimeoutExpired:

        return jsonify({
            "error": "Conversion timed out"
        }), 500


    except Exception as e:

        return jsonify({
            "error": str(e)
        }), 500


    finally:

        shutil.rmtree(
            job_folder,
            ignore_errors=True
        )


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
