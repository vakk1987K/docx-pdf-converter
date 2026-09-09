# =========================================================
# MULTIPLE IMAGES → PDF
# Existing routes are NOT changed.
# Endpoint: POST /images-to-pdf
# =========================================================

@app.route("/images-to-pdf", methods=["POST"])
def convert_multiple_images_to_pdf():

    import os
    import uuid
    import shutil

    from PIL import Image, ImageOps

    MAX_IMAGES = 20
    MAX_TOTAL_SIZE = 25 * 1024 * 1024

    allowed_extensions = {".jpg", ".jpeg", ".png"}

    job_id = uuid.uuid4().hex
    job_folder = os.path.join(BASE_FOLDER, f"multi-{job_id}")

    os.makedirs(job_folder, exist_ok=True)

    try:
        files = request.files.getlist("files")

        if not files:
            return jsonify({
                "success": False,
                "error": "Please select at least one JPG, JPEG or PNG image."
            }), 400

        if len(files) > MAX_IMAGES:
            return jsonify({
                "success": False,
                "error": f"Maximum {MAX_IMAGES} images are allowed."
            }), 400

        total_size = 0

        for uploaded_file in files:

            if not uploaded_file or not uploaded_file.filename:
                return jsonify({
                    "success": False,
                    "error": "One of the selected files is invalid."
                }), 400

            filename = secure_filename(uploaded_file.filename)

            if not filename:
                return jsonify({
                    "success": False,
                    "error": "Invalid filename detected."
                }), 400

            extension = os.path.splitext(filename)[1].lower()

            if extension not in allowed_extensions:
                return jsonify({
                    "success": False,
                    "error": "Only JPG, JPEG and PNG images are supported."
                }), 400

            # Check uploaded stream size without trusting client filename
            uploaded_file.stream.seek(0, os.SEEK_END)
            file_size = uploaded_file.stream.tell()
            uploaded_file.stream.seek(0)

            total_size += file_size

            if total_size > MAX_TOTAL_SIZE:
                return jsonify({
                    "success": False,
                    "error": "The total image size must not exceed 25 MB."
                }), 413

        converted_images = []

        for index, uploaded_file in enumerate(files):

            input_path = os.path.join(
                job_folder,
                f"input_{index}.upload"
            )

            uploaded_file.save(input_path)

            try:
                # Verify that the file is actually a valid image.
                with Image.open(input_path) as test_image:
                    test_image.verify()

            except Exception:
                return jsonify({
                    "success": False,
                    "error": f"Image {index + 1} is not a valid JPG, JPEG or PNG file."
                }), 400

            try:
                with Image.open(input_path) as image:

                    # Correct phone-camera orientation.
                    image = ImageOps.exif_transpose(image)

                    # Convert transparent / palette images safely.
                    if image.mode in ("RGBA", "LA", "P"):
                        background = Image.new(
                            "RGB",
                            image.size,
                            "white"
                        )

                        if image.mode == "P":
                            image = image.convert("RGBA")

                        background.paste(
                            image,
                            mask=image.getchannel("A")
                            if "A" in image.getbands()
                            else None
                        )

                        image = background

                    else:
                        image = image.convert("RGB")

                    output_image = os.path.join(
                        job_folder,
                        f"page_{index:03d}.jpg"
                    )

                    # Re-save as JPEG so every PDF page uses
                    # a predictable RGB image format.
                    image.save(
                        output_image,
                        "JPEG",
                        quality=92,
                        optimize=True
                    )

                    converted_images.append(output_image)

            except Exception as e:
                print("Image processing error:", str(e))

                return jsonify({
                    "success": False,
                    "error": f"Could not process image {index + 1}."
                }), 400

        if not converted_images:
            return jsonify({
                "success": False,
                "error": "No valid images were found."
            }), 400

        # -------------------------------------------------
        # Create ONE PDF containing all images
        # in the exact order received.
        # -------------------------------------------------

        pdf_path = os.path.join(
            job_folder,
            "multiple-images.pdf"
        )

        pdf_images = []

        try:

            for image_path in converted_images:
                image = Image.open(image_path)

                # Load into memory before closing file handle.
                image.load()

                pdf_images.append(image)

            first_image = pdf_images[0]
            remaining_images = pdf_images[1:]

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

        if not os.path.exists(pdf_path):
            raise RuntimeError("PDF file was not created.")

        if os.path.getsize(pdf_path) == 0:
            raise RuntimeError("Generated PDF is empty.")

        print(
            f"Multiple images converted successfully: "
            f"{len(converted_images)} images"
        )

        response = send_file(
            pdf_path,
            mimetype="application/pdf",
            as_attachment=True,
            download_name="multiple-images.pdf"
        )

        # Delete temporary files after response is finished.
        @response.call_on_close
        def cleanup():
            try:
                shutil.rmtree(job_folder, ignore_errors=True)
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
            "error": "Unable to create the PDF. Please try again."
        }), 500
