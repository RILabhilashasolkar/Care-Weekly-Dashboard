"""
Care Weekly Dashboard Portal – FastAPI backend
"""

from __future__ import annotations
import io
import json
import os
import sqlite3
import uuid
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from engine import analyze_file

# ---------------------------------------------------------------------------
# Directories & database
# ---------------------------------------------------------------------------

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = BASE_DIR / "uploads"
STATIC_DIR = BASE_DIR / "static"
DB_PATH = DATA_DIR / "reports.db"


def _db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def _init_db() -> None:
    DATA_DIR.mkdir(exist_ok=True)
    UPLOAD_DIR.mkdir(exist_ok=True)
    with _db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS reports (
                id          TEXT PRIMARY KEY,
                filename    TEXT NOT NULL,
                week_label  TEXT NOT NULL,
                created_at  TEXT NOT NULL,
                results     TEXT NOT NULL
            )
        """)
        conn.commit()


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------

app = FastAPI(title="Care Weekly Dashboard Portal", docs_url=None, redoc_url=None)


@app.on_event("startup")
def startup() -> None:
    _init_db()


# ---------------------------------------------------------------------------
# API routes
# ---------------------------------------------------------------------------

@app.post("/api/analyze")
async def api_analyze(
    file: UploadFile = File(...),
    week_label: str = Form(default=""),
):
    content = await file.read()
    if not content:
        raise HTTPException(400, "Uploaded file is empty.")
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Only Excel files (.xlsx / .xls) are supported.")

    try:
        results = analyze_file(io.BytesIO(content))
    except ValueError as e:
        raise HTTPException(422, str(e))
    except Exception as e:
        raise HTTPException(500, f"Analysis failed: {e}")

    # Persist
    report_id = str(uuid.uuid4())
    safe_label = week_label.strip() or datetime.now().strftime("Week of %d %b %Y")

    upload_path = UPLOAD_DIR / f"{report_id}_{file.filename}"
    upload_path.write_bytes(content)

    with _db() as conn:
        conn.execute(
            "INSERT INTO reports VALUES (?, ?, ?, ?, ?)",
            (report_id, file.filename, safe_label,
             datetime.now().isoformat(), json.dumps(results)),
        )
        conn.commit()

    return {"report_id": report_id, "week_label": safe_label, "results": results}


@app.get("/api/reports")
def api_list_reports():
    with _db() as conn:
        rows = conn.execute(
            "SELECT id, filename, week_label, created_at FROM reports ORDER BY created_at DESC"
        ).fetchall()
    return [dict(r) for r in rows]


@app.get("/api/reports/{report_id}")
def api_get_report(report_id: str):
    with _db() as conn:
        row = conn.execute(
            "SELECT * FROM reports WHERE id = ?", (report_id,)
        ).fetchone()
    if not row:
        raise HTTPException(404, "Report not found.")
    r = dict(row)
    r["results"] = json.loads(r["results"])
    return r


@app.delete("/api/reports/{report_id}")
def api_delete_report(report_id: str):
    with _db() as conn:
        conn.execute("DELETE FROM reports WHERE id = ?", (report_id,))
        conn.commit()
    return {"deleted": report_id}


# ---------------------------------------------------------------------------
# Serve frontend
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
def serve_index():
    return (STATIC_DIR / "index.html").read_text()


# Mount static assets (CSS, JS if any)
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
