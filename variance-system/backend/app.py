from flask import Flask, request, send_file, jsonify, render_template
import io
from processor import build_variance_workbook

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20 MB per upload


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/process", methods=["POST"])
def process():
    file1 = request.files.get("projection_file")
    file2 = request.files.get("dispatch_file")

    if not file1 or not file2:
        return jsonify({"error": "Both files are required."}), 400

    try:
        xlsx_bytes, report = build_variance_workbook(file1.read(), file2.read())
    except Exception as exc:
        return jsonify({"error": f"Failed to process files: {exc}"}), 500

    # Stash the report in a response header (base64-free, simple JSON) so the
    # frontend can show a match summary alongside the download.
    import json
    import base64
    report_b64 = base64.b64encode(json.dumps(report).encode()).decode()

    resp = send_file(
        io.BytesIO(xlsx_bytes),
        as_attachment=True,
        download_name="variance_report.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    resp.headers["X-Match-Report"] = report_b64
    resp.headers["Access-Control-Expose-Headers"] = "X-Match-Report"
    return resp


if __name__ == "__main__":
    app.run(debug=True, port=5000)
