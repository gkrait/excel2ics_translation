"""Web app: upload Excel timeplan, choose teacher(s), download ICS calendar(s)."""

import io
import tempfile
import zipfile
from pathlib import Path

from flask import Flask, request, send_file, jsonify

from excel2ics import extract_classes_for_teacher, classes_to_ics

app = Flask(__name__, static_folder="static", static_url_path="")
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024  # 32 MB

ALLOWED_EXTENSIONS = {"xlsx", "xls"}


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/")
def index():
    return send_file(Path(__file__).parent / "static" / "index.html")


@app.route("/api/convert", methods=["POST"])
def convert():
    """Accept Excel file + teacher code(s); return ICS or ZIP of ICS files."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    file = request.files["file"]
    if not file or file.filename == "":
        return jsonify({"error": "No file selected"}), 400
    if not allowed_file(file.filename):
        return jsonify({"error": "Only .xlsx and .xls files are allowed"}), 400

    teachers_raw = request.form.get("teachers", "").strip()
    if not teachers_raw:
        return jsonify({"error": "Enter at least one teacher code (e.g. RS7, RS4)"}), 400
    teachers = [t.strip() for t in teachers_raw.split(",") if t.strip()]
    if not teachers:
        return jsonify({"error": "Enter at least one teacher code"}), 400

    try:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            file.save(tmp.name)
            tmp_path = tmp.name

        results = {}
        for teacher in teachers:
            classes = extract_classes_for_teacher(tmp_path, teacher)
            ics_content = classes_to_ics(classes, teacher)
            results[teacher] = ics_content
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        Path(tmp_path).unlink(missing_ok=True)

    if len(teachers) == 1:
        teacher = teachers[0]
        buf = io.BytesIO(results[teacher].encode("utf-8"))
        filename = f"{teacher.lower().replace(' ', '_')}_calendar.ics"
        return send_file(
            buf,
            mimetype="text/calendar",
            as_attachment=True,
            download_name=filename,
        )

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for teacher, ics_content in results.items():
            filename = f"{teacher.lower().replace(' ', '_')}_calendar.ics"
            zf.writestr(filename, ics_content.encode("utf-8"))
    zip_buf.seek(0)
    return send_file(
        zip_buf,
        mimetype="application/zip",
        as_attachment=True,
        download_name="teacher_calendars.ics.zip",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5050, debug=True)
