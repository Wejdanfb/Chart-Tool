"""
MCH Nursing Chart Review Checklist - Monthly Consolidation Tool

Run locally:
    pip install -r requirements.txt
    python app.py
Then open http://localhost:5050 in a browser.

All Excel work (reading, formatting-preserving writes, formula generation) happens
here in Python via openpyxl - see consolidate.py. The browser is only the upload /
review / download interface.
"""

import os
import re
import json
import uuid
import fcntl
import contextlib
import shutil
import logging
from pathlib import Path

from flask import Flask, request, jsonify, send_file, render_template

from consolidate import analyze_file, consolidate, ValidationError

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("mch_tool")

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {".xlsx"}
MAX_CONTENT_LENGTH = 100 * 1024 * 1024  # 100 MB total, generous for ~17 files
DEFAULT_DOWNLOAD_NAME = "CHART_MCH_Consolidated.xlsx"
INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH

# Batch state (batch_id -> ordered list of file records) is kept on disk, one small JSON
# file per batch, instead of in a Python dict. Under gunicorn there can be several worker
# processes, each with its own memory; an in-memory dict meant a file uploaded through one
# worker was invisible to the next request handled by another worker, so uploaded files
# never showed up in the list. The filesystem is shared by all workers, so it is the
# single source of truth. Same interface as before: BATCHES.get(...) / BATCHES[id] = ...
_BATCH_ID_RE = re.compile(r"^[0-9a-fA-F-]{8,64}$")


def _valid_batch_id(batch_id) -> bool:
    return bool(batch_id) and bool(_BATCH_ID_RE.match(str(batch_id)))


class _BatchStore:
    def _dir(self, batch_id):
        if not _valid_batch_id(batch_id):
            raise ValueError("Invalid batch id.")
        return UPLOAD_DIR / batch_id

    def get(self, batch_id, default=None):
        if not _valid_batch_id(batch_id):
            return default
        try:
            return json.loads((self._dir(batch_id) / "batch.json").read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return default

    def __setitem__(self, batch_id, records):
        d = self._dir(batch_id)
        d.mkdir(parents=True, exist_ok=True)
        tmp = d / f"batch.json.{os.getpid()}.tmp"
        tmp.write_text(json.dumps(records), encoding="utf-8")
        os.replace(tmp, d / "batch.json")  # atomic

    @contextlib.contextmanager
    def lock(self, batch_id):
        """Cross-process lock so two simultaneous requests can't overwrite each other."""
        d = self._dir(batch_id)
        d.mkdir(parents=True, exist_ok=True)
        with open(d / ".lock", "w") as fh:
            fcntl.flock(fh, fcntl.LOCK_EX)
            try:
                yield
            finally:
                fcntl.flock(fh, fcntl.LOCK_UN)


BATCHES = _BatchStore()


# internal output filename (on disk in OUTPUT_DIR, never seen by the user) -> the
# download filename the user actually asked for. Stored as a small sidecar file next to
# the output so every worker process sees it. Entirely separate from the consolidation
# engine: consolidate() always writes to its own internally-chosen path and knows
# nothing about what the user wants the downloaded copy called.
class _DownloadNames:
    def __setitem__(self, internal_name, download_name):
        (OUTPUT_DIR / f"{internal_name}.name").write_text(download_name, encoding="utf-8")

    def get(self, internal_name, default=None):
        try:
            return (OUTPUT_DIR / f"{internal_name}.name").read_text(encoding="utf-8") or default
        except OSError:
            return default


DOWNLOAD_NAMES = _DownloadNames()


def _safe_name(filename: str) -> str:
    return "".join(c for c in filename if c.isalnum() or c in "._- ") or "file.xlsx"


def _sanitize_output_filename(raw: str):
    """User-entered output filename -> a safe base name (no extension), or None if
    nothing usable was entered (caller falls back to the default name in that case).
    Never touches file content - purely a display/download-name concern."""
    name = (raw or "").strip()
    if not name:
        return None
    name = re.sub(r"\.xlsx?$", "", name, flags=re.IGNORECASE)  # drop a typed extension
    name = INVALID_FILENAME_CHARS.sub("", name)  # \ / : * ? " < > |
    name = re.sub(r"\s+", " ", name).strip()
    name = name.rstrip(". ")  # Windows disallows trailing dots/spaces
    name = name[:150]  # keep it reasonable
    return name or None


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/upload", methods=["POST"])
def upload():
    """Accept one or more files, validate each, return per-file status.
    Files are appended to the batch in the order received, and order is preserved
    across multiple calls (e.g. if the user adds a few more files later) - this is
    what determines patient-column order in the final report."""
    batch_id = request.form.get("batch_id") or str(uuid.uuid4())
    if not _valid_batch_id(batch_id):
        return jsonify({"error": "Invalid batch id."}), 400
    batch_dir = UPLOAD_DIR / batch_id
    batch_dir.mkdir(parents=True, exist_ok=True)

    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No files received."}), 400

    results = []
    new_records = []
    for f in files:
        original_name = f.filename or "file.xlsx"
        ext = os.path.splitext(original_name)[1].lower()
        if ext not in ALLOWED_EXTENSIONS:
            results.append({
                "filename": original_name, "valid": False,
                "error": "Not an .xlsx file.",
            })
            continue

        file_id = str(uuid.uuid4())
        saved_name = f"{file_id}_{_safe_name(original_name)}"
        saved_path = batch_dir / saved_name
        f.save(str(saved_path))

        info = analyze_file(str(saved_path), original_name)
        record = {
            "file_id": file_id,
            "filename": original_name,
            "path": str(saved_path),
            "valid": info.valid,
            "error": info.error,
            "warnings": info.warnings,
            "unit_name": info.unit_name,
            "patient_count": info.patient_count,
        }
        new_records.append(record)
        results.append({k: v for k, v in record.items() if k != "path"})

    # Append to whatever is already stored for this batch (another request/worker may
    # have added files meanwhile), under a lock.
    with BATCHES.lock(batch_id):
        BATCHES[batch_id] = (BATCHES.get(batch_id) or []) + new_records
        order = [r["file_id"] for r in BATCHES.get(batch_id)]

    return jsonify({"batch_id": batch_id, "results": results, "order": order})


@app.route("/api/batch/<batch_id>", methods=["GET"])
def get_batch(batch_id):
    records = BATCHES.get(batch_id, [])
    return jsonify({
        "batch_id": batch_id,
        "files": [{k: v for k, v in r.items() if k != "path"} for r in records],
    })


@app.route("/api/batch/<batch_id>/remove/<file_id>", methods=["POST"])
def remove_file(batch_id, file_id):
    if not _valid_batch_id(batch_id):
        return jsonify({"files": []})
    with BATCHES.lock(batch_id):
        records = BATCHES.get(batch_id, [])
        BATCHES[batch_id] = [r for r in records if r["file_id"] != file_id]
        remaining = BATCHES.get(batch_id)
    return jsonify({"files": [{k: v for k, v in r.items() if k != "path"} for r in remaining]})


@app.route("/api/batch/<batch_id>/reorder", methods=["POST"])
def reorder(batch_id):
    """Accepts {"order": [file_id, file_id, ...]} and reorders the batch to match -
    used if the user drags files to rearrange them before generating."""
    if not _valid_batch_id(batch_id):
        return jsonify({"ok": False}), 400
    order = request.json.get("order", [])
    with BATCHES.lock(batch_id):
        records = {r["file_id"]: r for r in BATCHES.get(batch_id, [])}
        BATCHES[batch_id] = [records[fid] for fid in order if fid in records]
    return jsonify({"ok": True})


@app.route("/api/generate", methods=["POST"])
def generate():
    batch_id = request.json.get("batch_id")
    custom_filename = request.json.get("filename", "")
    if not _valid_batch_id(batch_id):
        return jsonify({"error": "No files in this batch."}), 400
    records = BATCHES.get(batch_id, [])
    if not records:
        return jsonify({"error": "No files in this batch."}), 400

    invalid = [r for r in records if not r["valid"]]
    if invalid:
        return jsonify({
            "error": "Cannot generate - some files failed validation.",
            "invalid_files": [r["filename"] for r in invalid],
        }), 400

    ordered = [(r["path"], r["filename"]) for r in records]
    # Internal on-disk filename - purely a storage detail, unrelated to what the user
    # wants the downloaded file called. consolidate() only ever sees this path.
    internal_name = f"CHART_MCH_Consolidated_{batch_id[:8]}.xlsx"
    output_path = OUTPUT_DIR / internal_name

    try:
        report = consolidate(ordered, str(output_path))
    except ValidationError as e:
        logger.warning("Generation blocked for batch %s: %s", batch_id, e)
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        logger.exception("Unexpected error generating batch %s", batch_id)
        return jsonify({"error": f"Unexpected error while generating the report: {e}"}), 500

    # The full diagnostic picture (unusual merges, text mismatches, multiple Unit/Ward
    # values, excluded empty columns, dropdown-source notes, etc.) is logged here for
    # anyone running the tool locally to check in this terminal window - but is
    # deliberately left out of the success response shown in the browser (see
    # app.js/showResult): a normal monthly run should end with a clean, simple "done"
    # screen, not a wall of internal detail that reads like something went wrong.
    for f in report["files"]:
        for w in f.get("warnings", []):
            logger.info("[%s] %s", f["filename"], w)
    if report.get("dissolved_unexpected_merges"):
        logger.info("Dissolved %d unexpected merged cell(s) in the base file.",
                     report["dissolved_unexpected_merges"])
    if not report.get("dropdown_source"):
        logger.info("No dropdown validation was carried over (base file had none).")
    elif report.get("dropdown_verified"):
        logger.info("Dropdown verified: '%s' resolves to %s",
                     report["dropdown_source"], report.get("dropdown_values"))
    else:
        logger.warning("Dropdown source '%s' could NOT be verified to resolve to real "
                        "values - check the base file's dropdown setup.",
                        report.get("dropdown_source"))

    sanitized = _sanitize_output_filename(custom_filename)
    download_name = f"{sanitized}.xlsx" if sanitized else DEFAULT_DOWNLOAD_NAME
    DOWNLOAD_NAMES[internal_name] = download_name

    report["download_url"] = f"/api/download/{internal_name}"
    report["download_filename"] = download_name
    return jsonify(report)


@app.route("/api/download/<name>", methods=["GET"])
def download(name):
    if name != os.path.basename(name) or not name.endswith(".xlsx"):
        return jsonify({"error": "File not found."}), 404
    path = OUTPUT_DIR / name
    if not path.exists():
        return jsonify({"error": "File not found."}), 404
    download_name = DOWNLOAD_NAMES.get(name, DEFAULT_DOWNLOAD_NAME)
    return send_file(str(path), as_attachment=True, download_name=download_name)


if __name__ == "__main__":
    print("MCH Consolidation Tool running at http://localhost:5050")
    app.run(host="127.0.0.1", port=int(os.environ.get("PORT", 5050)), debug=False)
