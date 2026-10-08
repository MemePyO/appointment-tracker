"""JSON-string API over backend.py for the web frontend (runs under Pyodide
and under desktop Python for tests)."""

import base64
import json
import os
import sys
import tempfile

import backend


def parse_days_json(text, max_day=None):
    try:
        days = backend.parse_days(text, max_day if max_day is None else int(max_day))
        return json.dumps({"ok": True, "days": days})
    except ValueError as exc:
        return json.dumps({"ok": False, "error": str(exc)})


def filter_names_json(names_json, prefix):
    names = json.loads(names_json)
    return json.dumps(backend.filter_names(names, prefix))


def validate_project_json(project_json):
    try:
        project = backend.validate_project(json.loads(project_json))
        return json.dumps({"ok": True, "project": project})
    except ValueError as exc:
        return json.dumps({"ok": False, "error": str(exc)})


def export_xlsx_b64(project_json):
    try:
        project = json.loads(project_json)
    except json.JSONDecodeError as exc:
        return json.dumps({"ok": False, "error": f"Invalid JSON: {exc}"})
    if sys.platform == "emscripten":
        path = "/tmp/appointments.xlsx"
    else:
        path = os.path.join(tempfile.gettempdir(), "appointments_bridge.xlsx")
    try:
        backend.export_excel(project, path)
    except ValueError as exc:
        return json.dumps({"ok": False, "error": str(exc)})
    with open(path, "rb") as fh:
        data = fh.read()
    try:
        os.remove(path)
    except OSError:
        pass
    return json.dumps({
        "ok": True,
        "filename": "appointments.xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "b64": base64.b64encode(data).decode("ascii"),
    })
