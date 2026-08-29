import io
import json
import shutil
import tempfile
import threading
import zipfile
from pathlib import Path
from urllib.parse import urlparse

from flask import Flask, jsonify, render_template, request, send_file, abort

import universal_dataset_engine as engine

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # URL text only; datasets are streamed by engine.

JOB_ROOT = Path(tempfile.gettempdir()) / "universal_dataset_web_jobs"
JOB_ROOT.mkdir(parents=True, exist_ok=True)

# The existing engine uses global output/cache folders.
# A lock makes each conversion isolated and prevents two web requests
# from changing those globals at the same time.
CONVERSION_LOCK = threading.Lock()


def validate_url(value: str) -> str:
    value = (value or "").strip().strip('"').strip("'")
    parsed = urlparse(value)

    if parsed.scheme.lower() not in ("http", "https") or not parsed.netloc:
        raise ValueError("Please enter a valid HTTP/HTTPS URL.")

    if len(value) > 4000:
        raise ValueError("URL is too long.")

    return value


def convert_url(url: str):
    job_dir = Path(tempfile.mkdtemp(prefix="job_", dir=JOB_ROOT))
    output_dir = job_dir / "downloaded_datasets"
    cache_dir = job_dir / ".dataset_cache"
    output_dir.mkdir()
    cache_dir.mkdir()

    old_output = engine.OUTPUT_DIR
    old_cache = engine.CACHE_DIR
    old_save_extra = engine.SAVE_EXTRA

    try:
        with CONVERSION_LOCK:
            engine.OUTPUT_DIR = output_dir
            engine.CACHE_DIR = cache_dir
            engine.SAVE_EXTRA = False

            # Dataset-page URLs (including Hugging Face) are handled by
            # the existing engine. Generic URLs go through process_url().
            engine.process_url(url)

        files = sorted(output_dir.glob("*.xlsx"))
        if not files:
            raise RuntimeError("No Excel file could be created from this URL.")

        return job_dir, files

    except Exception:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise

    finally:
        engine.OUTPUT_DIR = old_output
        engine.CACHE_DIR = old_cache
        engine.SAVE_EXTRA = old_save_extra


def zip_excel_files(files, zip_name="converted_excel_files.zip"):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
        for path in files:
            z.write(path, arcname=path.name)
    buffer.seek(0)
    return buffer, zip_name


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/api/convert")
def api_convert():
    url = request.form.get("url", "")
    try:
        url = validate_url(url)
        job_dir, files = convert_url(url)

        items = []
        for path in files:
            rel = path.relative_to(job_dir).as_posix()
            items.append({
                "name": path.name,
                "download_url": f"/download/{job_dir.name}/{rel}",
                "size": path.stat().st_size,
            })

        return jsonify({
            "success": True,
            "message": f"Converted {len(files)} Excel file(s).",
            "files": items,
            "download_all_url": f"/download-all/{job_dir.name}",
        })

    except Exception as exc:
        return jsonify({
            "success": False,
            "error": f"{type(exc).__name__}: {exc}",
        }), 400


@app.get("/download/<job_id>/<path:relative_path>")
def download_one(job_id, relative_path):
    job_dir = JOB_ROOT / job_id
    target = (job_dir / relative_path).resolve()

    if not job_dir.exists() or job_dir.resolve() not in target.parents:
        abort(404)

    if not target.is_file() or target.suffix.lower() != ".xlsx":
        abort(404)

    return send_file(target, as_attachment=True, download_name=target.name)


@app.get("/download-all/<job_id>")
def download_all(job_id):
    job_dir = JOB_ROOT / job_id
    if not job_dir.exists():
        abort(404)

    files = sorted(job_dir.joinpath("downloaded_datasets").glob("*.xlsx"))
    if not files:
        abort(404)

    buffer, name = zip_excel_files(files)
    return send_file(
        buffer,
        as_attachment=True,
        download_name=name,
        mimetype="application/zip",
    )


@app.get("/health")
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
