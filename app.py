```python
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

        # Create a private LibreOffice profile
        # so multiple requests do not conflict.
        lo_profile = os.path.join(
            job_folder,
            "lo-profile"
        )

        os.makedirs(
            lo_profile,
            exist_ok=True
        )

        # PDF is opened by LibreOffice Draw.
        # Explicitly export it as a DOCX file.
        command = [
            "libreoffice",
            "--headless",
            "--nologo",
            "--nodefault",
            "--nofirststartwizard",
            "--norestore",
            f"-env:UserInstallation=file://{lo_profile}",
            "--convert-to",
            "docx:\"Office Open XML Text\"",
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

        # Sometimes LibreOffice may return a
        # successful process but fail to create
        # the expected file.
        if not os.path.exists(output_file):

            stdout_text = result.stdout.decode(
                "utf-8",
                errors="ignore"
            )

            stderr_text = result.stderr.decode(
                "utf-8",
                errors="ignore"
            )

            details = (
                stderr_text.strip()
                or stdout_text.strip()
                or "LibreOffice did not create the DOCX file."
            )

            return jsonify({
                "error": "PDF to Word conversion failed.",
                "details": details
            }), 500

        download_name = (
            os.path.splitext(original_name)[0]
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
            "error": "Conversion timed out. The PDF may be too large or complex."
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
```
