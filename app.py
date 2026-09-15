import os
import threading
import uuid

from flask import (
    Flask,
    render_template,
    request,
    jsonify,
    send_file,
)

from image_finder import process_excel


app = Flask(__name__)

# 20 MB max upload — protects against accidental huge files
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

jobs = {}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/start", methods=["POST"])
def start_job():

    if "file" not in request.files:
        return jsonify({
            "error": "No Excel file uploaded."
        }), 400

    file = request.files["file"]

    if file.filename == "":
        return jsonify({
            "error": "No Excel file selected."
        }), 400

    job_id = str(uuid.uuid4())

    input_filename = file.filename
    input_path = os.path.join(
        UPLOAD_FOLDER,
        f"{job_id}_{input_filename}"
    )

    file.save(input_path)

    listing_language = request.form.get(
        "listing_language",
        "english"
    )

    translate_to_english = (
        request.form.get("translate_to_english") == "true"
    )

    brand_filter = request.form.get(
        "brand_filter",
        ""
    ).strip()

    site_filter = request.form.get(
        "site_filter",
        ""
    ).strip()

    jobs[job_id] = {
        "status": "starting",
        "logs": [],
        "result": None,
        "error": None,
    }

    thread = threading.Thread(
        target=run_job,
        args=(
            job_id,
            input_path,
            listing_language,
            translate_to_english,
            brand_filter,
            site_filter,
        ),
        daemon=True,
    )

    thread.start()

    return jsonify({"job_id": job_id})


def run_job(
    job_id,
    input_path,
    listing_language,
    translate_to_english,
    brand_filter,
    site_filter,
):

    def log(message):
        jobs[job_id]["logs"].append(str(message))

    try:

        jobs[job_id]["status"] = "running"

        log("Starting...")
        log(f"Language: {listing_language}")
        log(
            f"Translate to English: "
            f"{translate_to_english}"
        )
        log(
            f"Brand filter: "
            f"{brand_filter or 'Disabled'}"
        )
        log(
            f"Site filter: "
            f"{site_filter or 'Disabled'}"
        )
        log("")

        result = process_excel(
            input_file=input_path,
            output_dir=OUTPUT_FOLDER,
            listing_language=listing_language,
            translate_to_english=translate_to_english,
            brand_filter=brand_filter,
            site_filter=site_filter,
            log_func=log,
        )

        jobs[job_id]["result"] = result
        jobs[job_id]["status"] = "completed"

        log("")
        log(f"Success: {result}")

    except Exception as e:

        jobs[job_id]["status"] = "error"
        jobs[job_id]["error"] = str(e)

        log("")
        log(f"ERROR: {e}")


@app.route("/status/<job_id>")
def job_status(job_id):

    if job_id not in jobs:
        return jsonify({"error": "Job not found."}), 404

    job = jobs[job_id]

    return jsonify({
        "status": job["status"],
        "logs": job["logs"],
        "error": job["error"],
        "result": job["result"],
    })


@app.route("/download/<job_id>")
def download(job_id):

    if job_id not in jobs:
        return "Job not found.", 404

    result = jobs[job_id]["result"]

    if not result:
        return "File is not ready.", 404

    if not os.path.exists(result):
        return "Output file no longer exists.", 404

    return send_file(
        result,
        as_attachment=True,
        download_name=os.path.basename(result),
    )


from waitress import serve

if __name__ == "__main__":
    # Render (and most PaaS providers) inject PORT at runtime.
    port = int(os.environ.get("PORT", 5000))
    serve(app, host="0.0.0.0", port=port)
