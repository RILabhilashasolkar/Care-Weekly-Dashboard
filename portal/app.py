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
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
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


@app.get("/api/trends")
def api_get_trends():
    with _db() as conn:
        rows = conn.execute(
            "SELECT id, week_label, created_at, results FROM reports "
            "WHERE week_label NOT IN ('Test', 'test1') "
            "ORDER BY created_at ASC"
        ).fetchall()
    seen: dict = {}
    for r in rows:
        seen[r["week_label"]] = dict(r)
    result = []
    for wl, r in seen.items():
        try:
            res = json.loads(r["results"])
            fo  = res.get("final_output", {})
            kap = res.get("kapture", {})
            sap = res.get("sap_tickets", {})
            rdin = res.get("rdin_ageing", {}) or {}
            total = fo.get("total", 0)
            if total == 0 or total > 500_000:
                continue
            comp = fo.get("complaints", {}).get("total", 0)
            re   = fo.get("request_enquiry", {}).get("total", 0)
            comp_bd = fo.get("complaints", {}).get("breakdown", {})
            re_bd   = fo.get("request_enquiry", {}).get("breakdown", {})
            cats = ["Repair", "Demo & Installation", "PMS",
                    "Delivery related", "Refund", "Return",
                    "Invoice/Billing related", "Warranty"]
            result.append({
                "week_label":      wl,
                "created_at":      r["created_at"],
                "report_id":       r["id"],
                "total":           total,
                "complaints":      comp,
                "request_enquiry": re,
                "complaint_rate":  round(comp / total * 100, 2) if total else 0,
                "re_rate":         round(re / total * 100, 2) if total else 0,
                "kapture_total":   kap.get("total", 0),
                "sap_total":       sap.get("total", 0),
                "comp_breakdown":  {b: comp_bd.get(b, {}).get("count", 0) for b in cats},
                "re_breakdown":    {b: re_bd.get(b, {}).get("count", 0) for b in cats},
                "rdin_total":      rdin.get("total", 0),
                "rdin_avg_tat":    rdin.get("avg_tat_resolved"),
                "rdin_open":       rdin.get("open_no_tat", 0),
            })
        except Exception:
            pass
    return sorted(result, key=lambda x: x["created_at"])


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

    # Auto-backfill: re-analyse stored file if any computed data is missing
    kap = r["results"].get("kapture", {})
    sap = r["results"].get("sap_tickets", {})
    needs_backfill = (
        "rdin" not in kap
        or "re_top_subcats" not in kap
        or "re_top_subcats" not in sap
        or "top_subcats" not in r["results"].get("so_output", {})
        or "kpi" not in r["results"]
        or "rdin_ageing" not in r["results"]
    )
    if needs_backfill:
        upload_files = sorted(UPLOAD_DIR.glob(f"{report_id}_*"))
        if upload_files:
            try:
                fresh = analyze_file(io.BytesIO(upload_files[0].read_bytes()))
                r["results"] = fresh
                with _db() as conn:
                    conn.execute(
                        "UPDATE reports SET results = ? WHERE id = ?",
                        (json.dumps(fresh), report_id),
                    )
                    conn.commit()
            except Exception:
                pass

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

@app.get("/download-template")
def download_template():
    path = STATIC_DIR / "upload_template.xlsx"
    if not path.exists():
        raise HTTPException(404, "Template file not found.")
    return FileResponse(
        path=str(path),
        filename="Care_Weekly_Dashboard_Template.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/", response_class=HTMLResponse)
def serve_index():
    content = (STATIC_DIR / "index.html").read_text()
    return HTMLResponse(content=content, headers={
        "Cache-Control": "no-cache, no-store, must-revalidate",
        "Pragma": "no-cache",
    })


# Mount static assets (CSS, JS if any)
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
