import os
import uuid
import logging
import threading
import tempfile
import traceback
from urllib.parse import quote

from flask import Flask, request, send_file, render_template_string, redirect, url_for, abort
from waitress import serve

from image_finder import process_excel

# ---------------- Proxy configuration (from env vars) ----------------
TUNNEL_URL = os.environ.get("PROXY_TUNNEL_URL", "").strip()
PROXY_USER = os.environ.get("PROXY_USER", "").strip()
PROXY_PASS = os.environ.get("PROXY_PASS", "").strip()

if TUNNEL_URL and PROXY_USER and PROXY_PASS:
    user = quote(PROXY_USER, safe="")
    passwd = quote(PROXY_PASS, safe="")

    if TUNNEL_URL.startswith("tcp://"):
        hostport = TUNNEL_URL[len("tcp://"):].rstrip("/")
    else:
        hostport = TUNNEL_URL.replace("https://", "").replace("http://", "").rstrip("/")

    # HTTP proxy (Every Proxy HTTP mode) — handles DNS on the proxy side
    proxy_url = f"http://{user}:{passwd}@{hostport}"
    os.environ["HTTP_PROXY"] = proxy_url
    os.environ["HTTPS_PROXY"] = proxy_url
    os.environ["ALL_PROXY"] = proxy_url

# ---------------- logging to stdout (shows in Render logs) ----------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
log = logging.getLogger("app")

if TUNNEL_URL and PROXY_USER and PROXY_PASS:
    log.info("Proxy configured: http://%s@%s", PROXY_USER, TUNNEL_URL)
else:
    log.info("No proxy configured — running direct")

app = Flask(__name__)

JOBS = {}
JOBS_LOCK = threading.Lock()


# ---------------- HTML ----------------
UPLOAD_HTML = """
<!doctype html>
<html><head><title>Image Finder</title>
<style>
 body { font-family: system-ui, sans-serif; max-width: 720px; margin: 40px auto; padding: 0 16px; }
 input[type=file] { margin: 10px 0; }
 button { padding: 8px 16px; cursor: pointer; }
</style>
</head><body>
  <h1>Excel Article &rarr; Image Finder</h1>
  <form action="/process" method="post" enctype="multipart/form-data">
    <input type="file" name="file" accept=".xlsx,.xls" required>
    <br>
    <button type="submit">Process</button>
  </form>
</body></html>
"""

JOB_HTML = """
<!doctype html>
<html><head><title>Job {{ job_id }}</title>
<style>
 body { font-family: system-ui, sans-serif; max-width: 900px; margin: 40px auto; padding: 0 16px; }
 pre { background:#111; color:#0f0; padding:12px; border-radius:6px; max-height: 60vh; overflow:auto; font-size:12px; }
 .status { padding:6px 12px; border-radius:6px; display:inline-block; }
 .running { background:#fff3cd; }
 .done    { background:#d4edda; }
 .error   { background:#f8d7da; }
 button { padding: 8px 16px; cursor: pointer; margin-right: 8px; }
</style>
{% if job.status == 'running' %}<meta http-equiv="refresh" content="2">{% endif %}
</head><body>
  <h1>Job {{ job_id }}</h1>
  <div class="status {{ job.status }}">Status: {{ job.status }}</div>

  {% if job.status == 'done' %}
    <p>
      <a href="/download/{{ job_id }}"><button>Download Excel</button></a>
      <a href="/"><button>Upload Another File</button></a>
    </p>
  {% endif %}

  {% if job.error %}
    <p style="color:red"><b>Error:</b> {{ job.error }}</p>
    <p><a href="/"><button>Try Again</button></a></p>
  {% endif %}

  <h3>Log</h3>
  <pre>{{ log_text }}</pre>
</body></html>
"""


# ---------------- routes ----------------
@app.route("/")
def index():
    return render_template_string(UPLOAD_HTML)


@app.route("/process", methods=["POST"])
def process():
    if "file" not in request.files:
        return "No file part", 400
    file = request.files["file"]
    if file.filename == "":
        return "No selected file", 400

    job_id = uuid.uuid4().hex[:12]
    with JOBS_LOCK:
        JOBS[job_id] = {
            "status": "running",
            "logs": [],
            "file": None,
            "error": None,
        }

    temp_dir = tempfile.mkdtemp(prefix=f"job_{job_id}_")
    input_path = os.path.join(temp_dir, file.filename)
    output_dir = os.path.join(temp_dir, "output")
    os.makedirs(output_dir, exist_ok=True)
    file.save(input_path)

    log.info("Job %s started | file=%s | temp=%s", job_id, file.filename, temp_dir)

    t = threading.Thread(
        target=_run_job,
        args=(job_id, input_path, output_dir),
        daemon=True,
    )
    t.start()

    return redirect(url_for("job_view", job_id=job_id))


@app.route("/job/<job_id>")
def job_view(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            abort(404)
        log_text = "\n".join(job["logs"])
        job_view_data = {
            "status": job["status"],
            "error": job["error"],
        }
    return render_template_string(JOB_HTML, job_id=job_id, job=job_view_data, log_text=log_text)


@app.route("/download/<job_id>")
def download(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job or not job["file"]:
            abort(404)
        path = job["file"]
    return send_file(
        path,
        as_attachment=True,
        download_name=os.path.basename(path),
    )


# ---------------- worker ----------------
def _run_job(job_id, input_path, output_dir):
    def job_log(msg):
        msg = str(msg)
        with JOBS_LOCK:
            JOBS[job_id]["logs"].append(msg)
        log.info("[job %s] %s", job_id, msg)

    try:
        output_excel = process_excel(input_path, output_dir, log_func=job_log)
        with JOBS_LOCK:
            JOBS[job_id]["file"] = output_excel
            JOBS[job_id]["status"] = "done"
        log.info("Job %s finished: %s", job_id, output_excel)
    except Exception as e:
        tb = traceback.format_exc()
        log.error("Job %s failed: %s\n%s", job_id, e, tb)
        with JOBS_LOCK:
            JOBS[job_id]["status"] = "error"
            JOBS[job_id]["error"] = f"{type(e).__name__}: {e}"
            JOBS[job_id]["logs"].append(tb)


# ---------------- entrypoint ----------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    log.info("Starting Waitress on 0.0.0.0:%d", port)
    serve(app, host="0.0.0.0", port=port)
