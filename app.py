from __future__ import annotations

from datetime import datetime
import os
from flask import Flask, jsonify, render_template, request, send_file

from dashboard_builder import build_bootstrap_payload, build_dashboard_payload, build_dashboard_workbook


app = Flask(__name__)


def export_filename(report_date: str) -> str:
    safe_date = report_date.replace("-", "")
    timestamp = datetime.now().strftime("%H%M%S")
    return f"Eluru_Village_Issues_Dashboard_{safe_date}_{timestamp}.xlsx"


@app.get("/")
def index() -> str:
    return render_template("index.html")


@app.get("/healthz")
def healthcheck():
    return jsonify({"status": "ok"})


@app.get("/api/bootstrap")
def bootstrap():
    return jsonify(build_bootstrap_payload())


@app.post("/api/dashboard")
def dashboard():
    payload = request.get_json(silent=True) or {}
    records = payload.get("records") or []
    report_date = payload.get("reportDate")
    return jsonify(build_dashboard_payload(records, report_date))


@app.post("/api/export")
def export_dashboard():
    payload = request.get_json(silent=True) or {}
    records = payload.get("records") or []
    report_date = payload.get("reportDate") or datetime.now().date().isoformat()

    workbook_stream = build_dashboard_workbook(records, report_date)
    return send_file(
        workbook_stream,
        as_attachment=True,
        download_name=export_filename(report_date),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    host = os.getenv("HOST", "127.0.0.1")
    port = int(os.getenv("PORT", "5000"))
    debug = os.getenv("FLASK_DEBUG", "").lower() in {"1", "true", "yes", "on"}
    app.run(host=host, port=port, debug=debug)
