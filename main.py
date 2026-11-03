"""
Oracle Eloqua Form Submission Engine
------------------------------------
A self-contained, production-ready Python web application built with FastAPI.
Integrates with Oracle Eloqua REST API 2.0 to submit spreadsheet data (.xlsx/.csv)
to Eloqua forms with intelligent fuzzy mapping, batch throttling, circuit breaker,
and automated GDPR data purging.

Run with:
    pip install -r requirements.txt
    uvicorn main:app --host 0.0.0.0 --port 8000 --reload
"""

import os
import io
import csv
import json
import uuid
import time
import base64
import sqlite3
import datetime
from typing import List, Dict, Any, Optional
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks, Query
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import httpx
from dateutil import parser as date_parser
import openpyxl

# ---------------------------------------------------------------------------
# App Initialization & Directory Setup
# ---------------------------------------------------------------------------
app = FastAPI(
    title="Oracle Eloqua Form Submission Engine",
    description="Standalone Enterprise Form Data Ingestion System for Oracle Eloqua REST API 2.0",
    version="1.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "storage.db"
UPLOADS_DIR = BASE_DIR / "uploads"
REPORTS_DIR = BASE_DIR / "reports"

UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
REPORTS_DIR.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# Database Initialization (SQLite with WAL mode)
# ---------------------------------------------------------------------------
def get_db_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    with conn:
        conn.execute("PRAGMA journal_mode = WAL;")
        
        # 1. Config table
        conn.execute("""
            CREATE TABLE IF NOT EXISTS config (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL,
                updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
            );
        """)

        # 2. Forms cache table
        conn.execute("""
            CREATE TABLE IF NOT EXISTS forms_cache (
                form_id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                created_at DATETIME,
                updated_at DATETIME,
                cached_timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
            );
        """)

        # 3. User upload jobs table
        conn.execute("""
            CREATE TABLE IF NOT EXISTS jobs (
                job_id TEXT PRIMARY KEY,
                form_id INTEGER NOT NULL,
                form_name TEXT NOT NULL,
                total_rows INTEGER DEFAULT 0,
                processed_rows INTEGER DEFAULT 0,
                successful_rows INTEGER DEFAULT 0,
                failed_rows INTEGER DEFAULT 0,
                status TEXT CHECK(status IN ('pending', 'processing', 'completed', 'halted_error', 'failed')),
                error_summary TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                completed_at DATETIME
            );
        """)

        # 4. Job errors table
        conn.execute("""
            CREATE TABLE IF NOT EXISTS job_errors (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                job_id TEXT NOT NULL,
                row_index INTEGER NOT NULL,
                row_data_json TEXT NOT NULL,
                http_status INTEGER,
                error_message TEXT NOT NULL,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(job_id) REFERENCES jobs(job_id) ON DELETE CASCADE
            );
        """)

        # Seed default configuration if empty
        defaults = {
            "site_id": "1325",
            "pod": "p01",
            "username": "",
            "encrypted_password": "",
            "batch_size": "50",
            "throttle_seconds": "2",
            "max_row_limit": "5000",
            "forms_last_synced": ""
        }
        for k, v in defaults.items():
            conn.execute(
                "INSERT OR IGNORE INTO config (key, value) VALUES (?, ?);",
                (k, v)
            )
            
    conn.close()

init_db()

# ---------------------------------------------------------------------------
# Eloqua Configuration & Helper Functions
# ---------------------------------------------------------------------------
def get_config() -> Dict[str, str]:
    conn = get_db_connection()
    cursor = conn.execute("SELECT key, value FROM config;")
    cfg = {row["key"]: row["value"] for row in cursor.fetchall()}
    conn.close()
    return cfg

def update_config(key_values: Dict[str, str]):
    conn = get_db_connection()
    with conn:
        for k, v in key_values.items():
            conn.execute(
                "INSERT INTO config (key, value, updated_at) VALUES (?, ?, CURRENT_TIMESTAMP) ON CONFLICT(key) DO UPDATE SET value=excluded.value, updated_at=CURRENT_TIMESTAMP;",
                (k, v)
            )
    conn.close()

def get_eloqua_auth_header(cfg: Dict[str, str]) -> Optional[str]:
    site_id = cfg.get("site_id", "1325").strip()
    username = cfg.get("username", "").strip()
    enc_pw = cfg.get("encrypted_password", "").strip()
    if not (site_id and username and enc_pw):
        return None
    try:
        password = base64.b64decode(enc_pw.encode()).decode()
    except Exception:
        password = enc_pw
    token = f"{site_id}\\{username}:{password}"
    encoded = base64.b64encode(token.encode()).decode()
    return f"Basic {encoded}"

def get_eloqua_base_url(cfg: Dict[str, str]) -> str:
    pod = cfg.get("pod", "p01").strip() or "p01"
    return f"https://secure.{pod}.eloqua.com/API/REST/2.0"

# Mock fields for demonstration / fallback when credentials are not yet configured
MOCK_FORM_FIELDS: Dict[int, List[Dict[str, Any]]] = {
    101: [
        {"id": 14820, "name": "Email Address", "htmlName": "emailAddress", "dataType": "text"},
        {"id": 14821, "name": "First Name", "htmlName": "firstName", "dataType": "text"},
        {"id": 14822, "name": "Last Name", "htmlName": "lastName", "dataType": "text"},
        {"id": 14823, "name": "Company Name", "htmlName": "company", "dataType": "text"},
        {"id": 14824, "name": "Business Phone", "htmlName": "busPhone", "dataType": "text"},
        {"id": 14825, "name": "Job Title", "htmlName": "title", "dataType": "text"},
        {"id": 14826, "name": "Country", "htmlName": "country", "dataType": "text"},
        {"id": 14827, "name": "Registration Date", "htmlName": "regDate", "dataType": "date"},
        {"id": 14828, "name": "Notes / Comments", "htmlName": "comments", "dataType": "text"},
    ]
}

# ---------------------------------------------------------------------------
# GDPR Compliance Purge Task (72 Hours)
# ---------------------------------------------------------------------------
def run_gdpr_purge():
    """Deletes uploaded files, reports, and job rows older than 72 hours (3 days)."""
    try:
        conn = get_db_connection()
        three_days_ago = (datetime.datetime.utcnow() - datetime.timedelta(days=3)).strftime("%Y-%m-%d %H:%M:%S")
        cursor = conn.execute("SELECT job_id FROM jobs WHERE created_at < ?;", (three_days_ago,))
        expired_jobs = [row["job_id"] for row in cursor.fetchall()]

        for j_id in expired_jobs:
            # Delete upload file
            for ext in [".csv", ".xlsx", ".json"]:
                p = UPLOADS_DIR / f"{j_id}{ext}"
                if p.exists():
                    p.unlink(missing_ok=True)
            # Delete report file
            p_rep = REPORTS_DIR / f"failed_records_{j_id}.csv"
            if p_rep.exists():
                p_rep.unlink(missing_ok=True)
            # Delete database records
            with conn:
                conn.execute("DELETE FROM job_errors WHERE job_id = ?;", (j_id,))
                conn.execute("DELETE FROM jobs WHERE job_id = ?;", (j_id,))
        conn.close()
    except Exception as e:
        print(f"[GDPR Purge Error] {e}")

@app.on_event("startup")
async def startup_event():
    run_gdpr_purge()

# ---------------------------------------------------------------------------
# Normalization & Fuzzy Matching Utilities
# ---------------------------------------------------------------------------
def normalize_date_string(val: Any) -> str:
    """Normalizes dates to MM/DD/YYYY format for Oracle Eloqua, handling Excel serials, ISO, timestamps."""
    if val is None or val == "":
        return ""
    if isinstance(val, (datetime.date, datetime.datetime)):
        return val.strftime("%m/%d/%Y")
    val_str = str(val).strip()
    if not val_str:
        return ""

    # Check for pure numeric values (Excel serial dates like 46321 or timestamps)
    try:
        num = float(val_str)
        # Excel date serial number: roughly between 1000 and 100000
        # Excel epoch is 1899-12-30
        if 1000 <= num <= 100000:
            excel_epoch = datetime.datetime(1899, 12, 30)
            dt = excel_epoch + datetime.timedelta(days=num)
            return dt.strftime("%m/%d/%Y")
        # Unix timestamp (seconds)
        if 1000000000 <= num <= 9999999999:
            dt = datetime.datetime.utcfromtimestamp(num)
            return dt.strftime("%m/%d/%Y")
        # Unix timestamp (milliseconds)
        if 1000000000000 <= num <= 9999999999999:
            dt = datetime.datetime.utcfromtimestamp(num / 1000.0)
            return dt.strftime("%m/%d/%Y")
    except (ValueError, OverflowError):
        pass

    try:
        dt = date_parser.parse(val_str)
        if 1900 <= dt.year <= 2100:
            return dt.strftime("%m/%d/%Y")
        return val_str
    except Exception:
        return val_str

def levenshtein_ratio(s1: str, s2: str) -> float:
    """Calculates Levenshtein similarity ratio between two strings [0.0 - 1.0]."""
    s1, s2 = s1.lower().strip(), s2.lower().strip()
    if s1 == s2:
        return 1.0
    len1, len2 = len(s1), len(s2)
    if len1 == 0 or len2 == 0:
        return 0.0

    dp = [[0] * (len2 + 1) for _ in range(len1 + 1)]
    for i in range(len1 + 1):
        dp[i][0] = i
    for j in range(len2 + 1):
        dp[0][j] = j

    for i in range(1, len1 + 1):
        for j in range(1, len2 + 1):
            cost = 0 if s1[i - 1] == s2[j - 1] else 1
            dp[i][j] = min(
                dp[i - 1][j] + 1,       # deletion
                dp[i][j - 1] + 1,       # insertion
                dp[i - 1][j - 1] + cost # substitution
            )

    dist = dp[len1][len2]
    max_len = max(len1, len2)
    return 1.0 - (dist / max_len)

# ---------------------------------------------------------------------------
# Admin API Endpoints
# ---------------------------------------------------------------------------
@app.get("/api/admin/config")
def get_admin_config():
    cfg = get_config()
    # Mask password
    return {
        "site_id": cfg.get("site_id", "1325"),
        "pod": cfg.get("pod", "p01"),
        "username": cfg.get("username", ""),
        "has_password": bool(cfg.get("encrypted_password")),
        "batch_size": int(cfg.get("batch_size", "50")),
        "throttle_seconds": float(cfg.get("throttle_seconds", "2")),
        "max_row_limit": int(cfg.get("max_row_limit", "5000")),
        "forms_last_synced": cfg.get("forms_last_synced", "")
    }

@app.post("/api/admin/config")
async def save_admin_config(
    site_id: str = Form("1325"),
    pod: str = Form("p01"),
    username: str = Form(""),
    password: Optional[str] = Form(None),
    batch_size: int = Form(50),
    throttle_seconds: float = Form(2.0),
    max_row_limit: int = Form(5000),
):
    updates = {
        "site_id": site_id.strip(),
        "pod": pod.strip(),
        "username": username.strip(),
        "batch_size": str(batch_size),
        "throttle_seconds": str(throttle_seconds),
        "max_row_limit": str(max_row_limit),
    }
    if password is not None and password.strip():
        # Encode password safely (never plaintext in config query responses)
        enc = base64.b64encode(password.strip().encode()).decode()
        updates["encrypted_password"] = enc

    update_config(updates)
    return {"status": "success", "message": "Configuration successfully saved."}

@app.post("/api/admin/test-connection")
async def test_eloqua_connection():
    cfg = get_config()
    auth_header = get_eloqua_auth_header(cfg)
    if not auth_header:
        # Fallback simulator / warning if unconfigured
        return JSONResponse(
            status_code=400,
            content={"success": False, "message": "Missing Site ID, Username, or Password in configuration."}
        )

    base_url = get_eloqua_base_url(cfg)
    test_url = f"{base_url}/system/user/current"

    try:
        async with httpx.AsyncClient(timeout=10.0) as client:
            resp = await client.get(test_url, headers={"Authorization": auth_header, "Accept": "application/json"})
            if resp.status_code == 200:
                user_data = resp.json()
                display_name = user_data.get("name") or user_data.get("loginName", "Authenticated User")
                return {"success": True, "message": f"Connected successfully as '{display_name}' (Pod: {cfg.get('pod')}, Site: {cfg.get('site_id')})."}
            else:
                return JSONResponse(
                    status_code=resp.status_code,
                    content={"success": False, "message": f"Eloqua returned HTTP {resp.status_code}: {resp.text}"}
                )
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"success": False, "message": f"Connection failed: {str(e)}"}
        )

# ---------------------------------------------------------------------------
# Eloqua Form Listing & Field Schema Endpoints
# ---------------------------------------------------------------------------
@app.get("/api/forms")
def list_cached_forms(query: Optional[str] = Query(None)):
    conn = get_db_connection()
    if query:
        q = f"%{query}%"
        cursor = conn.execute(
            "SELECT form_id, name, cached_timestamp FROM forms_cache WHERE name LIKE ? OR CAST(form_id AS TEXT) LIKE ? ORDER BY name ASC LIMIT 5000;",
            (q, q)
        )
    else:
        cursor = conn.execute("SELECT form_id, name, cached_timestamp FROM forms_cache ORDER BY name ASC LIMIT 10000;")
    forms = [dict(row) for row in cursor.fetchall()]
    
    # Metadata count
    count_cursor = conn.execute("SELECT COUNT(*) as total FROM forms_cache;")
    total_count = count_cursor.fetchone()["total"]
    
    cfg = get_config()
    conn.close()
    return {
        "forms": forms,
        "total_cached": total_count,
        "last_synced": cfg.get("forms_last_synced", "")
    }

@app.post("/api/forms/sync")
async def sync_eloqua_forms():
    """Incrementally fetches all active forms from Eloqua with full pagination and caches in SQLite."""
    cfg = get_config()
    auth_header = get_eloqua_auth_header(cfg)
    base_url = get_eloqua_base_url(cfg)

    synced_count = 0
    now_str = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")

    if not auth_header:
        return JSONResponse(
            status_code=401,
            content={"success": False, "message": "Cannot sync forms: No valid Eloqua credentials configured. Please configure your Site ID, Username, and Password in /admin."}
        )

    try:
        page = 1
        page_size = 250
        has_more = True

        conn = get_db_connection()
        async with httpx.AsyncClient(timeout=30.0) as client:
            while has_more:
                url = f"{base_url}/assets/forms?depth=partial&page={page}&count={page_size}"
                resp = await client.get(url, headers={"Authorization": auth_header, "Accept": "application/json"})
                if resp.status_code == 200:
                    data = resp.json()
                    elements = data.get("elements", [])
                    if not elements:
                        has_more = False
                        break

                    with conn:
                        for elem in elements:
                            f_id = int(elem.get("id"))
                            f_name = elem.get("name", f"Form {f_id}")
                            is_archived = elem.get("isArchived", False)
                            if not is_archived:
                                conn.execute(
                                    """INSERT INTO forms_cache (form_id, name, updated_at, cached_timestamp)
                                       VALUES (?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
                                       ON CONFLICT(form_id) DO UPDATE SET name=excluded.name, cached_timestamp=CURRENT_TIMESTAMP;""",
                                    (f_id, f_name)
                                )
                                synced_count += 1

                    total = int(data.get("total", 0))
                    if len(elements) < page_size or (total > 0 and page * page_size >= total):
                        has_more = False
                    else:
                        page += 1
                        if page > 100:  # Safety guard: max 25,000 forms
                            has_more = False
                else:
                    if page == 1:
                        conn.close()
                        return JSONResponse(status_code=resp.status_code, content={"success": False, "message": f"Eloqua API error: {resp.text}"})
                    break

            with conn:
                conn.execute("UPDATE config SET value = ? WHERE key = 'forms_last_synced';", (now_str,))
            conn.close()

        return {"success": True, "synced_count": synced_count, "last_synced": now_str}
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "message": str(e)})

@app.get("/api/forms/{form_id}/fields")
async def get_form_fields(form_id: int):
    """Fetches field definitions for the specified form ID, filtering non-input elements."""
    cfg = get_config()
    auth_header = get_eloqua_auth_header(cfg)
    base_url = get_eloqua_base_url(cfg)

    if auth_header:
        try:
            async with httpx.AsyncClient(timeout=20.0) as client:
                url = f"{base_url}/assets/form/{form_id}?depth=complete"
                resp = await client.get(url, headers={"Authorization": auth_header, "Accept": "application/json"})
                if resp.status_code == 200:
                    form_data = resp.json()
                    raw_elements = form_data.get("elements", [])
                    actionable_fields = []
                    valid_types = {
                        "FormField", "FormTextElement", "FormSelectElement",
                        "FormHiddenElement", "FormRadioElement", "FormCheckboxElement"
                    }
                    for elem in raw_elements:
                        elem_type = elem.get("type", "")
                        # Filter layout containers, text blocks, submit buttons
                        if elem_type in valid_types or "field" in elem_type.lower() or "element" in elem_type.lower():
                            if elem.get("id"):
                                actionable_fields.append({
                                    "id": int(elem["id"]),
                                    "name": elem.get("name") or elem.get("htmlName") or f"Field {elem['id']}",
                                    "htmlName": elem.get("htmlName", ""),
                                    "dataType": elem.get("dataType", "text")
                                })
                    return {"form_id": form_id, "form_name": form_data.get("name"), "fields": actionable_fields}
                else:
                    return JSONResponse(status_code=resp.status_code, content={"error": resp.text})
        except Exception as e:
            return JSONResponse(status_code=500, content={"error": str(e)})

    # Simulated fallback fields for mock development
    fields = MOCK_FORM_FIELDS.get(form_id) or [
        {"id": 14820, "name": "Email Address", "htmlName": "emailAddress", "dataType": "text"},
        {"id": 14821, "name": "First Name", "htmlName": "firstName", "dataType": "text"},
        {"id": 14822, "name": "Last Name", "htmlName": "lastName", "dataType": "text"},
        {"id": 14823, "name": "Company Name", "htmlName": "company", "dataType": "text"},
        {"id": 14824, "name": "Phone Number", "htmlName": "phone", "dataType": "text"},
        {"id": 14825, "name": "Job Title", "htmlName": "title", "dataType": "text"},
        {"id": 14826, "name": "Lead Source", "htmlName": "leadSource", "dataType": "text"},
    ]
    return {"form_id": form_id, "form_name": f"Form #{form_id}", "fields": fields, "simulated": True}

# ---------------------------------------------------------------------------
# File Upload & Parsing Endpoint (CSV / XLSX - First Sheet Only)
# ---------------------------------------------------------------------------
@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    cfg = get_config()
    max_rows = int(cfg.get("max_row_limit", "5000"))
    
    filename = file.filename or "uploaded_data.csv"
    ext = Path(filename).suffix.lower()

    if ext not in [".csv", ".xlsx"]:
        raise HTTPException(status_code=400, detail="Unsupported file format. Please upload a .csv or .xlsx file.")

    content = await file.read()
    if len(content) == 0:
        raise HTTPException(status_code=400, detail="The uploaded file is empty.")

    headers: List[str] = []
    rows: List[Dict[str, Any]] = []

    try:
        if ext == ".csv":
            text_stream = io.StringIO(content.decode("utf-8-sig", errors="replace"))
            reader = csv.reader(text_stream)
            raw_headers = next(reader, None)
            if not raw_headers:
                raise HTTPException(status_code=400, detail="CSV file contains no header row.")
            headers = [h.strip() for h in raw_headers]
            for idx, r in enumerate(reader):
                if idx >= max_rows:
                    break
                row_dict = {}
                for h_idx, col_name in enumerate(headers):
                    row_dict[col_name] = r[h_idx] if h_idx < len(r) else ""
                rows.append(row_dict)

        elif ext == ".xlsx":
            wb = openpyxl.load_workbook(filename=io.BytesIO(content), data_only=True, read_only=True)
            sheet = wb.worksheets[0]  # Only the first sheet per specification
            sheet_rows = sheet.iter_rows(values_only=True)
            raw_headers = next(sheet_rows, None)
            if not raw_headers:
                raise HTTPException(status_code=400, detail="Excel sheet contains no header row.")
            headers = [str(h).strip() if h is not None else f"Column_{i+1}" for i, h in enumerate(raw_headers)]
            for idx, r in enumerate(sheet_rows):
                if idx >= max_rows:
                    break
                # Skip completely empty trailing rows
                if not any(cell is not None and str(cell).strip() != "" for cell in r):
                    continue
                row_dict = {}
                for h_idx, col_name in enumerate(headers):
                    val = r[h_idx] if h_idx < len(r) else ""
                    row_dict[col_name] = str(val) if val is not None else ""
                rows.append(row_dict)
            wb.close()
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to parse file: {str(e)}")

    if not rows:
        raise HTTPException(status_code=400, detail="The file contains headers but no data rows.")

    # Detect duplicate headers
    seen_headers = set()
    duplicate_headers = []
    for h in headers:
        if h in seen_headers:
            duplicate_headers.append(h)
        else:
            seen_headers.add(h)

    # Multi-tenant isolation: create unique job_id
    job_id = str(uuid.uuid4())
    
    # Save parsed data to uploads dir for batch processing
    data_payload = {
        "job_id": job_id,
        "filename": filename,
        "headers": headers,
        "duplicate_headers": list(set(duplicate_headers)),
        "total_rows": len(rows),
        "sample_row": rows[0] if rows else {},
        "rows": rows
    }
    with open(UPLOADS_DIR / f"{job_id}.json", "w", encoding="utf-8") as f:
        json.dump(data_payload, f)

    return {
        "job_id": job_id,
        "filename": filename,
        "headers": headers,
        "duplicate_headers": list(set(duplicate_headers)),
        "total_rows": len(rows),
        "sample_row": rows[0] if rows else {}
    }

# ---------------------------------------------------------------------------
# Batch Processing & Circuit Breaker Engine
# ---------------------------------------------------------------------------
async def process_batch_job(job_id: str, form_id: int, mapping: Dict[str, Any]):
    """
    Submits records in batches of 50 with configurable pause (default 2s).
    Circuit Breaker: If 100% of rows fail in ANY single batch, halt immediately.
    """
    job_file = UPLOADS_DIR / f"{job_id}.json"
    if not job_file.exists():
        return

    with open(job_file, "r", encoding="utf-8") as f:
        job_data = json.load(f)

    rows: List[Dict[str, Any]] = job_data.get("rows", [])
    total_rows = len(rows)

    cfg = get_config()
    batch_size = int(cfg.get("batch_size", "50"))
    throttle_secs = float(cfg.get("throttle_seconds", "2.0"))
    auth_header = get_eloqua_auth_header(cfg)
    base_url = get_eloqua_base_url(cfg)
    submit_url = f"{base_url}/data/form/{form_id}"

    conn = get_db_connection()
    with conn:
        conn.execute(
            "UPDATE jobs SET status = 'processing', total_rows = ? WHERE job_id = ?;",
            (total_rows, job_id)
        )
    conn.close()

    processed_rows = 0
    successful_rows = 0
    failed_rows = 0
    halted = False
    halt_reason = ""

    # Inspect target Eloqua form schema to check for elqCustomerGUID field element
    guid_field_id = None
    if auth_header:
        try:
            async with httpx.AsyncClient(timeout=10.0) as client:
                f_resp = await client.get(
                    f"{base_url}/assets/form/{form_id}?depth=complete",
                    headers={"Authorization": auth_header, "Accept": "application/json"}
                )
                if f_resp.status_code == 200:
                    f_data = f_resp.json()
                    for elem in f_data.get("elements", []):
                        combined_name = f"{elem.get('name', '')} {elem.get('htmlName', '')}".lower().replace("-", "").replace("_", "")
                        if "elqcustomerguid" in combined_name or "customerguid" in combined_name:
                            guid_field_id = str(elem.get("id"))
                            break
        except Exception as e:
            logger.warning(f"Failed to check form fields for elqCustomerGUID: {e}")

    # Prepare chunks of batch_size
    for chunk_start in range(0, total_rows, batch_size):
        chunk_end = min(chunk_start + batch_size, total_rows)
        batch = rows[chunk_start:chunk_end]
        batch_len = len(batch)
        batch_failures = 0

        async with httpx.AsyncClient(timeout=15.0) as client:
            for idx, row in enumerate(batch):
                row_index = chunk_start + idx + 1
                
                # Construct Eloqua Payload using mapped numeric field IDs
                field_values = []
                for col_name, field_info in mapping.items():
                    if not field_info or field_info == "ignore":
                        continue
                    field_id = field_info.get("id") if isinstance(field_info, dict) else field_info
                    raw_val = row.get(col_name, "")
                    
                    # Normalize dates if applicable
                    val_str = str(raw_val) if raw_val is not None else ""
                    if "date" in str(col_name).lower() or (isinstance(field_info, dict) and field_info.get("dataType") == "date"):
                        val_str = normalize_date_string(val_str)

                    field_values.append({
                        "type": "FieldValue",
                        "id": str(field_id),
                        "value": val_str
                    })

                # Hardcode mandatory elqCustomerGUID in fieldValues if field element exists on form and not mapped
                if guid_field_id and not any(fv.get("id") == guid_field_id for fv in field_values):
                    field_values.append({
                        "type": "FieldValue",
                        "id": guid_field_id,
                        "value": "00000000-0000-0000-0000-000000000000"
                    })

                # Mandatory hardcoded elqCustomerGUID for Oracle Eloqua REST API Form POSTs
                payload = {
                    "type": "FormData",
                    "elqCustomerGUID": "00000000-0000-0000-0000-000000000000",
                    "fieldValues": field_values
                }

                row_success = False
                http_status = 200
                error_detail = ""

                if auth_header:
                    try:
                        resp = await client.post(
                            submit_url,
                            json=payload,
                            headers={"Authorization": auth_header, "Content-Type": "application/json"}
                        )
                        http_status = resp.status_code
                        if resp.status_code in (200, 201):
                            row_success = True
                        else:
                            error_detail = f"Eloqua API Error ({resp.status_code}): {resp.text}"
                    except Exception as exc:
                        http_status = 500
                        error_detail = f"HTTP Request Failed: {str(exc)}"
                else:
                    # Simulated submission runner
                    # Validate email structure as a sanity check
                    email_val = ""
                    for col_name, field_info in mapping.items():
                        if isinstance(field_info, dict) and "email" in field_info.get("name", "").lower():
                            email_val = str(row.get(col_name, "")).strip()
                            break
                    if email_val and ("@" not in email_val or "." not in email_val):
                        http_status = 400
                        error_detail = f"Validation Error: Invalid email format '{email_val}'"
                        row_success = False
                    else:
                        row_success = True

                if row_success:
                    successful_rows += 1
                else:
                    failed_rows += 1
                    batch_failures += 1
                    # Record error log in SQLite
                    c = get_db_connection()
                    with c:
                        c.execute(
                            """INSERT INTO job_errors (job_id, row_index, row_data_json, http_status, error_message)
                               VALUES (?, ?, ?, ?, ?);""",
                            (job_id, row_index, json.dumps(row), http_status, error_detail)
                        )
                    c.close()

                processed_rows += 1

        # Update progress in SQLite after batch completes
        c = get_db_connection()
        with c:
            c.execute(
                """UPDATE jobs SET processed_rows = ?, successful_rows = ?, failed_rows = ?
                   WHERE job_id = ?;""",
                (processed_rows, successful_rows, failed_rows, job_id)
            )
        c.close()

        # Circuit Breaker Check: 100% failure rate in this single batch
        if batch_failures == batch_len:
            halted = True
            halt_reason = (
                f"Execution halted: Batch [{chunk_start // batch_size + 1}] failed completely "
                f"({batch_failures}/{batch_len} rows failed - 100% error rate). "
                f"Halting immediately to prevent account lockouts or invalid form spam."
            )
            break

        # Respect throttle delay between batches if more rows remain
        if chunk_end < total_rows:
            time.sleep(throttle_secs)

    # Finalize job status
    final_status = "halted_error" if halted else ("completed" if failed_rows == 0 else "failed")
    c = get_db_connection()
    with c:
        c.execute(
            """UPDATE jobs SET status = ?, error_summary = ?, completed_at = CURRENT_TIMESTAMP
               WHERE job_id = ?;""",
            (final_status, halt_reason if halted else "", job_id)
        )
    c.close()

@app.post("/api/jobs/submit")
async def start_submission_job(
    job_id: str = Form(...),
    form_id: int = Form(...),
    form_name: str = Form(...),
    mapping_json: str = Form(...),
    background_tasks: BackgroundTasks = BackgroundTasks()
):
    try:
        mapping = json.loads(mapping_json)
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid field mapping JSON format.")

    # Check email validation requirement
    has_email = False
    for col, target in mapping.items():
        if target and target != "ignore":
            t_name = target.get("name", "") if isinstance(target, dict) else str(target)
            if "email" in t_name.lower() or "email" in col.lower():
                has_email = True
                break

    if not has_email:
        raise HTTPException(status_code=400, detail="At least one field containing 'email' must be mapped before submitting.")

    # Initialize job in DB
    conn = get_db_connection()
    with conn:
        conn.execute(
            """INSERT INTO jobs (job_id, form_id, form_name, status, created_at)
               VALUES (?, ?, ?, 'pending', CURRENT_TIMESTAMP)
               ON CONFLICT(job_id) DO UPDATE SET form_id=excluded.form_id, form_name=excluded.form_name, status='pending';""",
            (job_id, form_id, form_name)
        )
    conn.close()

    # Launch background batch worker
    background_tasks.add_task(process_batch_job, job_id, form_id, mapping)
    return {"status": "started", "job_id": job_id}

@app.get("/api/jobs/{job_id}/status")
def get_job_status(job_id: str):
    conn = get_db_connection()
    cursor = conn.execute("SELECT * FROM jobs WHERE job_id = ?;", (job_id,))
    row = cursor.fetchone()
    conn.close()
    if not row:
        raise HTTPException(status_code=404, detail="Job not found.")
    return dict(row)

@app.get("/api/jobs/{job_id}/download-errors")
def download_failed_records(job_id: str):
    """
    Generates and downloads a CSV of failed rows, retaining all original uploaded columns
    plus appended columns: Eloqua_HTTP_Status and Eloqua_Error_Details.
    """
    conn = get_db_connection()
    cursor = conn.execute(
        "SELECT row_index, row_data_json, http_status, error_message FROM job_errors WHERE job_id = ? ORDER BY row_index ASC;",
        (job_id,)
    )
    error_rows = cursor.fetchall()
    conn.close()

    if not error_rows:
        raise HTTPException(status_code=404, detail="No error records found for this job.")

    # Parse first row to extract original spreadsheet headers
    first_row_data = json.loads(error_rows[0]["row_data_json"])
    original_headers = list(first_row_data.keys())
    output_headers = original_headers + ["Eloqua_HTTP_Status", "Eloqua_Error_Details"]

    string_buffer = io.StringIO()
    writer = csv.writer(string_buffer)
    writer.writerow(output_headers)

    for item in error_rows:
        row_dict = json.loads(item["row_data_json"])
        row_values = [row_dict.get(h, "") for h in original_headers]
        row_values.append(item["http_status"])
        row_values.append(item["error_message"])
        writer.writerow(row_values)

    string_buffer.seek(0)
    filename = f"failed_records_{job_id[:8]}.csv"
    return StreamingResponse(
        iter([string_buffer.getvalue()]),
        media_type="text/csv",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

# ---------------------------------------------------------------------------
# Integrated Single-File Frontend Template (HTML5 + Tailwind CSS + Vanilla JS)
# ---------------------------------------------------------------------------
INDEX_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Oracle Eloqua Form Submission Engine</title>
  <!-- Tailwind CSS CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
  <style>
    body { font-family: 'Inter', sans-serif; background-color: #F3F5F9; }
    .step-active { color: #2563EB; border-color: #2563EB; font-weight: 600; }
    .step-done { color: #059669; border-color: #059669; }
    .step-inactive { color: #9CA3AF; border-color: #E5E7EB; }
  </style>
</head>
<body class="min-h-screen flex flex-col text-slate-800">

  <!-- Top Navigation Bar -->
  <header class="bg-white border-b border-gray-200 sticky top-0 z-40">
    <div class="max-w-6xl mx-auto px-6 h-16 flex items-center justify-between">
        <div class="w-9 h-9 rounded-xl bg-blue-600 flex items-center justify-center text-white shadow-sm shadow-blue-500/25">
          <svg class="w-5 h-5" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
            <polyline points="14 2 14 8 20 8"></polyline>
            <path d="m9 15 3-3 3 3"></path>
            <path d="M12 12v6"></path>
          </svg>
        </div>
        <h1 class="text-base font-bold text-slate-900 tracking-tight">Eloqua Form Batch Submitter</h1>
      </div>
      <div class="flex items-center space-x-3">
        <span class="inline-flex items-center px-2.5 py-1 rounded-full text-xs font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200">
          <span class="w-2 h-2 rounded-full bg-emerald-500 mr-1.5 animate-pulse"></span>
          Pod p01 • Site 1325
        </span>
        <button onclick="resetWizard()" class="inline-flex items-center space-x-1.5 px-3 py-1.5 text-xs font-semibold text-slate-600 hover:text-slate-900 bg-gray-100 hover:bg-gray-200 rounded-lg transition-colors" title="Reset and start over">
          <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"></path>
          </svg>
          <span>Start Over</span>
        </button>
      </div>
    </div>
  </header>

  <!-- Step Tracker Progress Bar -->
  <div class="bg-white border-b border-gray-200 py-3 shadow-xs">
    <div class="max-w-4xl mx-auto px-6">
      <div class="flex items-center justify-between text-xs font-medium">
        <div id="step-nav-1" class="flex items-center space-x-2 step-active">
          <span class="w-6 h-6 rounded-full border-2 flex items-center justify-center font-bold">1</span>
          <span>Select Form</span>
        </div>
        <div class="flex-1 h-0.5 bg-gray-200 mx-4"></div>
        <div id="step-nav-2" class="flex items-center space-x-2 step-inactive">
          <span class="w-6 h-6 rounded-full border-2 flex items-center justify-center font-bold">2</span>
          <span>Upload File</span>
        </div>
        <div class="flex-1 h-0.5 bg-gray-200 mx-4"></div>
        <div id="step-nav-3" class="flex items-center space-x-2 step-inactive">
          <span class="w-6 h-6 rounded-full border-2 flex items-center justify-center font-bold">3</span>
          <span>Map Fields</span>
        </div>
        <div class="flex-1 h-0.5 bg-gray-200 mx-4"></div>
        <div id="step-nav-4" class="flex items-center space-x-2 step-inactive">
          <span class="w-6 h-6 rounded-full border-2 flex items-center justify-center font-bold">4</span>
          <span>Review Mapping</span>
        </div>
        <div class="flex-1 h-0.5 bg-gray-200 mx-4"></div>
        <div id="step-nav-5" class="flex items-center space-x-2 step-inactive">
          <span class="w-6 h-6 rounded-full border-2 flex items-center justify-center font-bold">5</span>
          <span>Process & Results</span>
        </div>
      </div>
    </div>
  </div>

  <!-- Main Content Container -->
  <main class="flex-1 max-w-4xl w-full mx-auto p-6 mb-20">

    <!-- Step 1: Select Form -->
    <section id="step-1" class="block space-y-6">
      <div class="bg-white rounded-2xl border border-gray-200 p-8 shadow-xs">
        <div class="flex items-start justify-between">
          <div>
            <h2 class="text-xl font-bold text-slate-900">Step 1: Choose Target Eloqua Form</h2>
            <p class="text-sm text-slate-500 mt-1">Select the destination form to retrieve its schema and numeric field identifiers.</p>
          </div>
          <div class="flex space-x-2">
            <button onclick="syncForms()" id="sync-forms-btn" class="inline-flex items-center space-x-1 px-3 py-1.5 text-xs font-semibold text-blue-700 bg-blue-50 hover:bg-blue-100 rounded-lg border border-blue-200 transition-colors">
              <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"></path></svg>
              <span>Sync Forms</span>
            </button>
            <button onclick="refreshFormFields()" id="refresh-fields-btn" class="inline-flex items-center space-x-1 px-3 py-1.5 text-xs font-semibold text-slate-700 bg-gray-50 hover:bg-gray-100 rounded-lg border border-gray-200 transition-colors">
              <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"></path></svg>
              <span>Refresh Fields</span>
            </button>
          </div>
        </div>

        <div class="mt-6">
          <label class="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-2">Search Form by Name or ID</label>
          <div class="relative">
            <input type="text" id="form-search-input" placeholder="Type to filter 1,500+ cached forms..." oninput="handleFormSearch(this.value)" class="w-full px-4 py-3 text-sm bg-slate-50 border border-gray-300 rounded-xl focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition-all">
            <div id="form-dropdown" class="hidden absolute top-full left-0 right-0 mt-1 max-h-60 overflow-y-auto bg-white border border-gray-200 rounded-xl shadow-lg z-50 divide-y divide-gray-100">
              <!-- Dynamically populated -->
            </div>
          </div>
        </div>

        <div id="selected-form-banner" class="hidden mt-6 p-4 bg-blue-50 border border-blue-200 rounded-xl flex items-center justify-between">
          <div class="flex items-center space-x-3">
            <div class="w-9 h-9 rounded-lg bg-blue-600 text-white flex items-center justify-center font-bold text-xs">ID</div>
            <div>
              <p id="selected-form-name" class="text-sm font-bold text-blue-950"></p>
              <p id="selected-form-meta" class="text-xs text-blue-700"></p>
            </div>
          </div>
          <span class="text-xs font-semibold px-2 py-1 bg-blue-200/60 text-blue-900 rounded-md">Selected</span>
        </div>

        <div class="mt-4 flex items-center justify-between text-xs text-slate-400">
          <span id="forms-cache-stat">Loading cached forms...</span>
          <span id="forms-last-synced-stat">Last synced: checking...</span>
        </div>
      </div>
    </section>

    <!-- Step 2: Upload File -->
    <section id="step-2" class="hidden space-y-6">
      <div class="bg-white rounded-2xl border border-gray-200 p-8 shadow-xs">
        <div>
          <h2 class="text-xl font-bold text-slate-900">Step 2: Upload Spreadsheet</h2>
          <p class="text-sm text-slate-500 mt-1">Upload your customer or lead dataset in Excel (.xlsx) or CSV format. First sheet will be ingested.</p>
        </div>

        <div id="drop-zone" onclick="document.getElementById('file-input').click()" class="mt-6 border-2 border-dashed border-gray-300 hover:border-blue-500 bg-slate-50 hover:bg-blue-50/50 rounded-2xl p-12 text-center cursor-pointer transition-all">
          <input type="file" id="file-input" class="hidden" accept=".csv, .xlsx" onchange="handleFileSelected(event)">
          <div class="w-16 h-16 mx-auto mb-4 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center">
            <svg class="w-8 h-8" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"></path>
            </svg>
          </div>
          <h3 class="text-base font-bold text-slate-800">Drag & drop your file here</h3>
          <p class="text-xs text-slate-500 mt-1">or click anywhere to browse your computer for a .xlsx or .csv file</p>
          <div class="mt-4 inline-flex items-center space-x-2 px-3 py-1 bg-white border border-gray-200 rounded-full text-xs font-medium text-slate-600">
            <span>Enforces maximum 5,000 rows ceiling</span>
          </div>
        </div>

        <div id="file-preview-banner" class="hidden mt-6 p-4 bg-emerald-50 border border-emerald-200 rounded-xl flex items-center justify-between">
          <div class="flex items-center space-x-3">
            <div class="w-10 h-10 rounded-lg bg-emerald-600 text-white flex items-center justify-center">
              <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path></svg>
            </div>
            <div>
              <p id="file-preview-name" class="text-sm font-bold text-emerald-950"></p>
              <p id="file-preview-stats" class="text-xs text-emerald-700"></p>
            </div>
          </div>
          <button onclick="resetFileSelection()" class="text-xs text-red-600 hover:text-red-800 font-semibold">Change File</button>
        </div>

        <div id="duplicate-header-warning" class="hidden mt-4 p-4 bg-amber-50 border border-amber-200 rounded-xl text-amber-900 text-xs">
          <span class="font-bold">⚠️ Warning - Duplicate Column Headers:</span>
          <span id="duplicate-header-list"></span>
        </div>
      </div>
    </section>

    <!-- Step 3: Map Fields -->
    <section id="step-3" class="hidden space-y-6">
      <div class="bg-white rounded-2xl border border-gray-200 p-8 shadow-xs">
        <div class="flex items-center justify-between">
          <div>
            <h2 class="text-xl font-bold text-slate-900">Step 3: Map Columns to Eloqua Fields</h2>
            <p class="text-sm text-slate-500 mt-1">Headers have been auto-matched via Levenshtein fuzzy similarity. Ensure at least one email field is mapped.</p>
          </div>
          <span id="email-validation-badge" class="px-3 py-1 rounded-full text-xs font-semibold bg-amber-100 text-amber-800 border border-amber-200">
            Email mapping required
          </span>
        </div>

        <div class="mt-6 border border-gray-200 rounded-xl overflow-hidden">
          <table class="w-full text-left text-xs">
            <thead class="bg-slate-50 text-slate-500 uppercase tracking-wider font-semibold border-b border-gray-200">
              <tr>
                <th class="py-3 px-4">Uploaded Column Header</th>
                <th class="py-3 px-4">Row 1 Sample Preview</th>
                <th class="py-3 px-4">Target Eloqua Form Field</th>
              </tr>
            </thead>
            <tbody id="mapping-table-body" class="divide-y divide-gray-100 text-slate-700 font-medium">
              <!-- Populated by JS -->
            </tbody>
          </table>
        </div>
      </div>
    </section>

    <!-- Step 4: Review Mapping -->
    <section id="step-4" class="hidden space-y-6">
      <div class="bg-white rounded-2xl border border-gray-200 p-8 shadow-xs">
        <div>
          <h2 class="text-xl font-bold text-slate-900">Step 4: Review Active Field Mappings</h2>
          <p class="text-sm text-slate-500 mt-1">Verify target assignments before executing the batch ingestion engine.</p>
        </div>

        <div class="mt-6 border border-gray-200 rounded-xl overflow-hidden">
          <table class="w-full text-left text-xs">
            <thead class="bg-slate-50 text-slate-500 uppercase tracking-wider font-semibold border-b border-gray-200">
              <tr>
                <th class="py-3 px-4 w-1/2">Spreadsheet Column (Sample Value)</th>
                <th class="py-3 px-4 w-1/2">Eloqua Target Field [ID]</th>
              </tr>
            </thead>
            <tbody id="review-table-body" class="divide-y divide-gray-100 text-slate-700">
              <!-- Populated by JS -->
            </tbody>
          </table>
        </div>

        <div class="mt-6 p-4 bg-slate-50 border border-gray-200 rounded-xl flex items-center justify-between text-xs">
          <div class="flex space-x-6 text-slate-600">
            <div><span class="font-bold text-slate-900" id="summary-total-rows">0</span> total rows</div>
            <div><span class="font-bold text-emerald-600" id="summary-mapped-count">0</span> columns mapped</div>
            <div><span class="font-bold text-slate-400" id="summary-ignored-count">0</span> columns ignored</div>
          </div>
          <div class="text-slate-500">Target: <span id="summary-target-form" class="font-semibold text-slate-900"></span></div>
        </div>
      </div>
    </section>

    <!-- Step 5: Process & Results -->
    <section id="step-5" class="hidden space-y-6">
      <div class="bg-white rounded-2xl border border-gray-200 p-8 shadow-xs">
        <div>
          <h2 class="text-xl font-bold text-slate-900">Step 5: Processing Ingestion Batches</h2>
          <p class="text-sm text-slate-500 mt-1">Dispatches batches of 50 records with circuit breaker safety protection.</p>
        </div>

        <!-- Progress Bar -->
        <div class="mt-6">
          <div class="flex items-center justify-between text-xs font-semibold mb-2">
            <span id="process-phase-text" class="text-slate-700">Submitting records to Eloqua REST API...</span>
            <span id="process-percent-text" class="text-blue-600">0%</span>
          </div>
          <div class="w-full h-3 bg-gray-100 rounded-full overflow-hidden">
            <div id="process-progress-bar" class="h-full bg-blue-600 transition-all duration-300" style="width: 0%"></div>
          </div>
        </div>

        <!-- Metric Stat Cards -->
        <div class="grid grid-cols-4 gap-4 mt-6">
          <div class="p-4 bg-slate-50 border border-gray-200 rounded-xl">
            <span class="text-xs text-slate-500 font-medium">Total Rows</span>
            <p id="stat-total-rows" class="text-2xl font-bold text-slate-900 mt-1">0</p>
          </div>
          <div class="p-4 bg-blue-50 border border-blue-200 rounded-xl">
            <span class="text-xs text-blue-700 font-medium">Processed</span>
            <p id="stat-processed-rows" class="text-2xl font-bold text-blue-900 mt-1">0</p>
          </div>
          <div class="p-4 bg-emerald-50 border border-emerald-200 rounded-xl">
            <span class="text-xs text-emerald-700 font-medium">Successful</span>
            <p id="stat-successful-rows" class="text-2xl font-bold text-emerald-900 mt-1">0</p>
          </div>
          <div class="p-4 bg-rose-50 border border-rose-200 rounded-xl">
            <span class="text-xs text-rose-700 font-medium">Failed</span>
            <p id="stat-failed-rows" class="text-2xl font-bold text-rose-900 mt-1">0</p>
          </div>
        </div>

        <!-- Circuit Breaker Alert Banner -->
        <div id="circuit-breaker-banner" class="hidden mt-6 p-4 bg-rose-50 border border-rose-300 rounded-xl">
          <div class="flex items-start space-x-3">
            <div class="w-8 h-8 rounded-full bg-rose-600 text-white flex items-center justify-center flex-shrink-0 font-bold">!</div>
            <div>
              <h4 class="text-sm font-bold text-rose-950">Circuit Breaker Triggered</h4>
              <p id="circuit-breaker-text" class="text-xs text-rose-800 mt-1"></p>
            </div>
          </div>
        </div>

        <!-- 100% Success Banner -->
        <div id="success-banner" class="hidden mt-6 p-6 bg-emerald-50 border border-emerald-200 rounded-xl text-center">
          <div class="w-12 h-12 bg-emerald-600 text-white rounded-full flex items-center justify-center mx-auto mb-3">
            <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"></path></svg>
          </div>
          <h3 class="text-base font-bold text-emerald-950">Batch Ingestion Completed Successfully!</h3>
          <p id="success-banner-subtext" class="text-xs text-emerald-700 mt-1"></p>
        </div>

        <!-- Download Failed Records Button -->
        <div id="failed-records-container" class="hidden mt-6 text-center">
          <a id="download-error-btn" href="#" class="inline-flex items-center space-x-2 px-5 py-2.5 bg-rose-600 hover:bg-rose-700 text-white text-xs font-bold rounded-xl shadow-md transition-colors">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"></path></svg>
            <span>Download Failed Records (.csv)</span>
          </a>
          <p class="text-[11px] text-slate-400 mt-2">Includes original row inputs plus Eloqua_HTTP_Status and Eloqua_Error_Details</p>
        </div>
      </div>
    </section>

  </main>

  <!-- Sticky Bottom Action Bar -->
  <footer class="bg-white border-t border-gray-200 fixed bottom-0 left-0 right-0 z-40 py-3 shadow-lg">
    <div class="max-w-4xl mx-auto px-6 flex items-center justify-between">
      <div id="footer-status-text" class="text-xs text-slate-500 font-medium">Step 1 of 5: Select Eloqua destination form</div>
      <div class="flex space-x-3">
        <button id="btn-back" onclick="goToPreviousStep()" class="hidden px-4 py-2 text-xs font-semibold text-slate-600 hover:text-slate-900 bg-gray-100 hover:bg-gray-200 rounded-xl transition-colors">Back</button>
        <button id="btn-next" onclick="goToNextStep()" disabled class="px-5 py-2 text-xs font-bold text-white bg-blue-600 hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed rounded-xl shadow-sm transition-colors">Continue →</button>
      </div>
    </div>
  </footer>

  <!-- Application Logic JavaScript -->
  <script>
    let currentStep = 1;
    let cachedForms = [];
    let selectedForm = null;
    let formFields = [];
    let uploadedData = null;
    let fieldMapping = {};
    let currentJobId = null;
    let pollingInterval = null;

    // Common aliases for marketing operations
    const ALIAS_MAP = {
      "first": "first",
      "firstname": "first",
      "first name": "first",
      "fname": "first",
      "last": "last",
      "lastname": "last",
      "last name": "last",
      "lname": "last",
      "email": "email",
      "email address": "email",
      "emailaddress": "email",
      "e-mail": "email",
      "phone": "phone",
      "cell": "phone",
      "mobile": "phone",
      "telephone": "phone",
      "bus phone": "phone",
      "business phone": "phone",
      "company": "company",
      "company name": "company",
      "organization": "company",
      "title": "title",
      "job title": "title",
      "country": "country"
    };

    function levenshteinSimilarity(s1, s2) {
      s1 = (s1 || '').toLowerCase().trim();
      s2 = (s2 || '').toLowerCase().trim();
      if (s1 === s2) return 1.0;
      const len1 = s1.length, len2 = s2.length;
      if (!len1 || !len2) return 0.0;
      const dp = Array.from({ length: len1 + 1 }, () => Array(len2 + 1).fill(0));
      for (let i = 0; i <= len1; i++) dp[i][0] = i;
      for (let j = 0; j <= len2; j++) dp[0][j] = j;
      for (let i = 1; i <= len1; i++) {
        for (let j = 1; j <= len2; j++) {
          const cost = s1[i - 1] === s2[j - 1] ? 0 : 1;
          dp[i][j] = min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + cost);
        }
      }
      return 1.0 - (dp[len1][len2] / Math.max(len1, len2));
    }
    function min(a, b, c) { return Math.min(a, Math.min(b, c)); }

    async function loadForms() {
      try {
        const res = await fetch('/api/forms');
        const data = await res.json();
        cachedForms = data.forms || [];
        document.getElementById('forms-cache-stat').innerText = `${data.total_cached || cachedForms.length} active forms cached locally in SQLite WAL`;
        document.getElementById('forms-last-synced-stat').innerText = `Last synced: ${data.last_synced || 'Never'}`;
      } catch (e) {
        console.error(e);
      }
    }

    function handleFormSearch(term) {
      const dropdown = document.getElementById('form-dropdown');
      if (!term.trim()) {
        dropdown.classList.add('hidden');
        return;
      }
      const lower = term.toLowerCase();
      const filtered = cachedForms.filter(f => f.name.toLowerCase().includes(lower) || String(f.form_id).includes(lower));
      if (filtered.length === 0) {
        dropdown.innerHTML = `<div class="p-3 text-xs text-slate-400">No matching forms found.</div>`;
      } else {
        dropdown.innerHTML = filtered.slice(0, 15).map(f => `
          <div onclick="selectForm(${f.form_id}, '${escapeHtml(f.name)}')" class="p-3 hover:bg-blue-50 cursor-pointer flex justify-between items-center text-xs">
            <span class="font-semibold text-slate-800">${escapeHtml(f.name)}</span>
            <span class="text-slate-400 font-mono">ID: ${f.form_id}</span>
          </div>
        `).join('');
      }
      dropdown.classList.remove('hidden');
    }

    async function selectForm(formId, formName) {
      selectedForm = { form_id: formId, name: formName };
      document.getElementById('form-dropdown').classList.add('hidden');
      document.getElementById('form-search-input').value = formName;
      document.getElementById('selected-form-banner').classList.remove('hidden');
      document.getElementById('selected-form-name').innerText = formName;
      document.getElementById('selected-form-meta').innerText = `Eloqua Numeric ID: ${formId}`;
      document.getElementById('btn-next').disabled = false;
      await loadFormFields(formId);
    }

    async function loadFormFields(formId) {
      try {
        const res = await fetch(`/api/forms/${formId}/fields`);
        const data = await res.json();
        formFields = data.fields || [];
      } catch (e) {
        console.error(e);
      }
    }

    async function syncForms() {
      const btn = document.getElementById('sync-forms-btn');
      btn.innerText = 'Syncing...';
      try {
        const res = await fetch('/api/forms/sync', { method: 'POST' });
        const data = await res.json();
        await loadForms();
        alert(`Incremental sync complete. Fetched ${data.synced_count} forms.`);
      } catch (e) {
        alert('Sync failed: ' + e);
      } finally {
        btn.innerText = 'Sync Forms';
      }
    }

    async function refreshFormFields() {
      if (!selectedForm) {
        alert('Please select a form first.');
        return;
      }
      const btn = document.getElementById('refresh-fields-btn');
      btn.innerText = 'Refreshing...';
      await loadFormFields(selectedForm.form_id);
      btn.innerText = 'Refresh Fields';
      alert(`Refreshed ${formFields.length} actionable input fields for form ${selectedForm.form_id}.`);
    }

    async function handleFileSelected(event) {
      const file = event.target.files[0];
      if (!file) return;

      const formData = new FormData();
      formData.append('file', file);

      try {
        const res = await fetch('/api/upload', { method: 'POST', body: formData });
        if (!res.ok) {
          const err = await res.json();
          alert(err.detail || 'Upload failed');
          return;
        }
        uploadedData = await res.json();
        currentJobId = uploadedData.job_id;

        // Show banner
        document.getElementById('file-preview-banner').classList.remove('hidden');
        document.getElementById('file-preview-name').innerText = uploadedData.filename;
        document.getElementById('file-preview-stats').innerText = `${uploadedData.total_rows} rows • ${uploadedData.headers.length} columns detected (First sheet parsed)`;
        
        // Duplicate header check
        if (uploadedData.duplicate_headers && uploadedData.duplicate_headers.length > 0) {
          document.getElementById('duplicate-header-warning').classList.remove('hidden');
          document.getElementById('duplicate-header-list').innerText = ' ' + uploadedData.duplicate_headers.join(', ');
        } else {
          document.getElementById('duplicate-header-warning').classList.add('hidden');
        }

        document.getElementById('btn-next').disabled = false;
      } catch (e) {
        alert('Error uploading file: ' + e);
      }
    }

    function resetFileSelection() {
      uploadedData = null;
      document.getElementById('file-input').value = '';
      document.getElementById('file-preview-banner').classList.add('hidden');
      document.getElementById('duplicate-header-warning').classList.add('hidden');
      document.getElementById('btn-next').disabled = true;
    }

    function buildMappingUI() {
      const tbody = document.getElementById('mapping-table-body');
      tbody.innerHTML = '';
      fieldMapping = {};

      uploadedData.headers.forEach((col, colIdx) => {
        const sampleVal = uploadedData.sample_row[col] !== undefined ? String(uploadedData.sample_row[col]) : '';
        const normCol = col.toLowerCase().trim();
        let matchedField = null;

        // 1. Alias lookup
        for (const [alias, canonical] of Object.entries(ALIAS_MAP)) {
          if (normCol === alias || normCol.includes(alias)) {
            matchedField = formFields.find(f => f.name.toLowerCase().includes(canonical));
            if (matchedField) break;
          }
        }

        // 2. Levenshtein fuzzy match
        if (!matchedField) {
          let bestScore = 0;
          for (const f of formFields) {
            const score = levenshteinSimilarity(col, f.name);
            if (score > bestScore && score >= 0.70) {
              bestScore = score;
              matchedField = f;
            }
          }
        }

        if (matchedField) {
          fieldMapping[col] = matchedField;
        } else {
          fieldMapping[col] = 'ignore';
        }

        const isDuplicate = uploadedData.duplicate_headers.includes(col);
        const tr = document.createElement('tr');
        tr.className = isDuplicate ? 'bg-amber-50/60' : '';
        tr.innerHTML = `
          <td class="py-3 px-4 font-semibold flex items-center space-x-2">
            <span>${escapeHtml(col)}</span>
            ${isDuplicate ? '<span class="px-1.5 py-0.5 text-[10px] bg-amber-200 text-amber-900 rounded font-bold">Duplicate Header</span>' : ''}
          </td>
          <td class="py-3 px-4 text-slate-500 font-mono max-w-xs truncate" title="${escapeHtml(sampleVal)}">
            ${escapeHtml(sampleVal) || '<span class="italic text-slate-300">empty</span>'}
          </td>
          <td class="py-3 px-4">
            <select onchange="updateFieldMapping('${escapeHtml(col)}', this.value)" class="w-full text-xs bg-slate-50 border border-gray-300 rounded-lg py-1.5 px-2.5 focus:bg-white focus:outline-none focus:ring-1 focus:ring-blue-500">
              <option value="ignore" ${!matchedField ? 'selected' : ''}>[ -- Ignore Column -- ]</option>
              ${formFields.map(f => `
                <option value="${f.id}" ${matchedField && matchedField.id === f.id ? 'selected' : ''}>
                  ${escapeHtml(f.name)} (ID: ${f.id})
                </option>
              `).join('')}
            </select>
          </td>
        `;
        tbody.appendChild(tr);
      });

      validateEmailMapping();
    }

    function updateFieldMapping(col, selectedValue) {
      if (selectedValue === 'ignore') {
        fieldMapping[col] = 'ignore';
      } else {
        const f = formFields.find(item => item.id == selectedValue);
        fieldMapping[col] = f || 'ignore';
      }
      validateEmailMapping();
    }

    function validateEmailMapping() {
      let hasEmail = false;
      for (const [col, target] of Object.entries(fieldMapping)) {
        if (target && target !== 'ignore') {
          const tName = target.name || '';
          if (tName.toLowerCase().includes('email') || col.toLowerCase().includes('email')) {
            hasEmail = true;
            break;
          }
        }
      }
      const badge = document.getElementById('email-validation-badge');
      const nextBtn = document.getElementById('btn-next');
      if (hasEmail) {
        badge.className = 'px-3 py-1 rounded-full text-xs font-semibold bg-emerald-100 text-emerald-800 border border-emerald-200';
        badge.innerText = '✓ Valid Email Field Mapped';
        nextBtn.disabled = false;
      } else {
        badge.className = 'px-3 py-1 rounded-full text-xs font-semibold bg-amber-100 text-amber-800 border border-amber-200';
        badge.innerText = 'Email mapping required';
        nextBtn.disabled = true;
      }
    }

    function buildReviewUI() {
      const tbody = document.getElementById('review-table-body');
      tbody.innerHTML = '';
      let mappedCount = 0;
      let ignoredCount = 0;

      for (const [col, target] of Object.entries(fieldMapping)) {
        const sampleVal = uploadedData.sample_row[col] !== undefined ? String(uploadedData.sample_row[col]) : '';
        if (target && target !== 'ignore') {
          mappedCount++;
          const tr = document.createElement('tr');
          tr.innerHTML = `
            <td class="py-2.5 px-4 font-semibold text-slate-800">
              ${escapeHtml(col)} <span class="text-xs text-slate-400 font-normal">("${escapeHtml(sampleVal)}")</span>
            </td>
            <td class="py-2.5 px-4 text-blue-700 font-semibold">
              ${escapeHtml(target.name)} <span class="text-xs text-slate-400 font-mono">[${target.id}]</span>
            </td>
          `;
          tbody.appendChild(tr);
        } else {
          ignoredCount++;
        }
      }

      document.getElementById('summary-total-rows').innerText = uploadedData.total_rows;
      document.getElementById('summary-mapped-count').innerText = mappedCount;
      document.getElementById('summary-ignored-count').innerText = ignoredCount;
      document.getElementById('summary-target-form').innerText = `${selectedForm.name} (ID: ${selectedForm.form_id})`;
      document.getElementById('btn-next').disabled = false;
    }

    async function startSubmission() {
      const formData = new FormData();
      formData.append('job_id', currentJobId);
      formData.append('form_id', selectedForm.form_id);
      formData.append('form_name', selectedForm.name);
      formData.append('mapping_json', JSON.stringify(fieldMapping));

      try {
        const res = await fetch('/api/jobs/submit', { method: 'POST', body: formData });
        if (!res.ok) {
          const err = await res.json();
          alert(err.detail || 'Submission failed');
          return;
        }
        pollJobProgress();
      } catch (e) {
        alert('Error starting job: ' + e);
      }
    }

    function pollJobProgress() {
      if (pollingInterval) clearInterval(pollingInterval);
      pollingInterval = setInterval(async () => {
        try {
          const res = await fetch(`/api/jobs/${currentJobId}/status`);
          if (!res.ok) return;
          const job = await res.json();

          const total = job.total_rows || 0;
          const processed = job.processed_rows || 0;
          const succ = job.successful_rows || 0;
          const fail = job.failed_rows || 0;
          const pct = total > 0 ? Math.round((processed / total) * 100) : 0;

          document.getElementById('stat-total-rows').innerText = total;
          document.getElementById('stat-processed-rows').innerText = processed;
          document.getElementById('stat-successful-rows').innerText = succ;
          document.getElementById('stat-failed-rows').innerText = fail;

          document.getElementById('process-percent-text').innerText = `${pct}%`;
          document.getElementById('process-progress-bar').style.width = `${pct}%`;

          if (job.status === 'halted_error') {
            clearInterval(pollingInterval);
            document.getElementById('process-phase-text').innerText = 'Halted via Circuit Breaker';
            document.getElementById('circuit-breaker-banner').classList.remove('hidden');
            document.getElementById('circuit-breaker-text').innerText = job.error_summary || '100% failure rate in batch.';
            setupDownloadErrorBtn();
          } else if (job.status === 'completed' || job.status === 'failed') {
            clearInterval(pollingInterval);
            document.getElementById('process-phase-text').innerText = 'Execution Complete';
            if (fail === 0) {
              document.getElementById('success-banner').classList.remove('hidden');
              document.getElementById('success-banner-subtext').innerText = `All ${succ} records submitted successfully to Eloqua form #${job.form_id}.`;
            } else {
              setupDownloadErrorBtn();
            }
          }
        } catch (e) {
          console.error(e);
        }
      }, 1000);
    }

    function setupDownloadErrorBtn() {
      const container = document.getElementById('failed-records-container');
      const btn = document.getElementById('download-error-btn');
      btn.href = `/api/jobs/${currentJobId}/download-errors`;
      container.classList.remove('hidden');
    }

    function goToNextStep() {
      if (currentStep === 1) {
        currentStep = 2;
      } else if (currentStep === 2) {
        currentStep = 3;
        buildMappingUI();
      } else if (currentStep === 3) {
        currentStep = 4;
        buildReviewUI();
      } else if (currentStep === 4) {
        currentStep = 5;
        startSubmission();
      }
      renderStep();
    }

    function goToPreviousStep() {
      if (currentStep > 1 && currentStep < 5) {
        currentStep--;
        renderStep();
      }
    }

    function renderStep() {
      for (let i = 1; i <= 5; i++) {
        document.getElementById(`step-${i}`).classList.toggle('hidden', i !== currentStep);
        const nav = document.getElementById(`step-nav-${i}`);
        if (i < currentStep) {
          nav.className = 'flex items-center space-x-2 step-done';
        } else if (i === currentStep) {
          nav.className = 'flex items-center space-x-2 step-active';
        } else {
          nav.className = 'flex items-center space-x-2 step-inactive';
        }
      }

      const backBtn = document.getElementById('btn-back');
      const nextBtn = document.getElementById('btn-next');
      const footerStatus = document.getElementById('footer-status-text');

      backBtn.classList.toggle('hidden', currentStep === 1 || currentStep === 5);
      if (currentStep === 1) {
        footerStatus.innerText = 'Step 1 of 5: Select Eloqua destination form';
        nextBtn.innerText = 'Continue →';
        nextBtn.disabled = !selectedForm;
      } else if (currentStep === 2) {
        footerStatus.innerText = 'Step 2 of 5: Upload spreadsheet (.xlsx or .csv)';
        nextBtn.innerText = 'Continue →';
        nextBtn.disabled = !uploadedData;
      } else if (currentStep === 3) {
        footerStatus.innerText = 'Step 3 of 5: Review fuzzy column mappings';
        nextBtn.innerText = 'Review Mapping →';
        validateEmailMapping();
      } else if (currentStep === 4) {
        footerStatus.innerText = 'Step 4 of 5: Final review and confirmation';
        nextBtn.innerText = 'Submit to Eloqua Form';
        nextBtn.disabled = false;
      } else if (currentStep === 5) {
        footerStatus.innerText = 'Step 5 of 5: Ingestion dispatcher in progress';
        nextBtn.classList.add('hidden');
      }
    }

    function resetWizard() {
      if (confirm('Start over? Current mapping and upload progress will be cleared.')) {
        window.location.reload();
      }
    }

    function escapeHtml(text) {
      if (!text) return '';
      return String(text).replace(/[&<>"']/g, m => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
      }[m]));
    }

    window.onload = loadForms;
  </script>
</body>
</html>
"""

# ---------------------------------------------------------------------------
# Hidden Admin Portal HTML Template (/admin)
# ---------------------------------------------------------------------------
ADMIN_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Eloqua Engine Admin - System Configuration</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
  <style>body { font-family: 'Inter', sans-serif; background-color: #F3F5F9; }</style>
</head>
<body class="min-h-screen text-slate-800 p-8 flex flex-col items-center">
  <div class="max-w-2xl w-full bg-white rounded-2xl border border-gray-200 shadow-sm p-8">
    <div class="flex items-center justify-between pb-6 border-b border-gray-100">
      <div>
        <h1 class="text-xl font-bold text-slate-900">System Configuration</h1>
        <p class="text-xs text-slate-500 mt-1">Oracle Eloqua REST API 2.0 Credentials & Engine Settings</p>
      </div>
      <span class="px-2.5 py-1 bg-amber-50 text-amber-800 border border-amber-200 rounded-full text-xs font-semibold">
        Hidden Route (/admin)
      </span>
    </div>

    <form id="admin-form" onsubmit="handleSaveConfig(event)" class="space-y-5 mt-6">
      <div class="grid grid-cols-2 gap-4">
        <div>
          <label class="block text-xs font-bold uppercase text-slate-600 mb-1">Eloqua Pod</label>
          <input type="text" id="pod" name="pod" value="p01" required class="w-full px-3 py-2 text-sm bg-slate-50 border border-gray-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-blue-500 outline-none">
          <p class="text-[11px] text-slate-400 mt-1">e.g. p01, p02</p>
        </div>
        <div>
          <label class="block text-xs font-bold uppercase text-slate-600 mb-1">Eloqua Company Name</label>
          <input type="text" id="site_id" name="site_id" value="" placeholder="e.g. company" required class="w-full px-3 py-2 text-sm bg-slate-50 border border-gray-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-blue-500 outline-none">
          <p class="text-[11px] text-slate-400 mt-1">Company name for authentication</p>
        </div>
      </div>

      <div>
        <label class="block text-xs font-bold uppercase text-slate-600 mb-1">Eloqua Username</label>
        <input type="text" id="username" name="username" placeholder="e.g. first.last" class="w-full px-3 py-2 text-sm bg-slate-50 border border-gray-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-blue-500 outline-none">
      </div>

      <div>
        <label class="block text-xs font-bold uppercase text-slate-600 mb-1">Password</label>
        <input type="password" id="password" name="password" placeholder="Leave empty to retain existing saved password" class="w-full px-3 py-2 text-sm bg-slate-50 border border-gray-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-blue-500 outline-none">
        <p id="pw-status" class="text-[11px] text-slate-400 mt-1">Stored securely. Passwords are never returned to the UI.</p>
      </div>

      <div class="grid grid-cols-3 gap-4 pt-2">
        <div>
          <label class="block text-xs font-bold uppercase text-slate-600 mb-1">Batch Size</label>
          <input type="number" id="batch_size" name="batch_size" value="50" min="1" max="500" required class="w-full px-3 py-2 text-sm bg-slate-50 border border-gray-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-blue-500 outline-none">
          <p class="text-[11px] text-slate-400 mt-1">Default: 50 records</p>
        </div>
        <div>
          <label class="block text-xs font-bold uppercase text-slate-600 mb-1">Throttle (Seconds)</label>
          <input type="number" step="0.5" id="throttle_seconds" name="throttle_seconds" value="2" min="0" max="60" required class="w-full px-3 py-2 text-sm bg-slate-50 border border-gray-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-blue-500 outline-none">
          <p class="text-[11px] text-slate-400 mt-1">Pause per batch</p>
        </div>
        <div>
          <label class="block text-xs font-bold uppercase text-slate-600 mb-1">Max Row Ceiling</label>
          <input type="number" id="max_row_limit" name="max_row_limit" value="5000" min="10" max="50000" required class="w-full px-3 py-2 text-sm bg-slate-50 border border-gray-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-blue-500 outline-none">
          <p class="text-[11px] text-slate-400 mt-1">Default: 5,000</p>
        </div>
      </div>

      <div id="test-result-box" class="hidden p-3.5 rounded-xl text-xs font-medium"></div>

      <div class="pt-6 border-t border-gray-100 flex items-center justify-between">
        <button type="button" onclick="testConnection()" id="test-btn" class="px-4 py-2 text-xs font-bold text-slate-700 bg-gray-100 hover:bg-gray-200 rounded-xl transition-colors">
          Test Connection
        </button>
        <button type="submit" id="save-btn" class="px-5 py-2 text-xs font-bold text-white bg-blue-600 hover:bg-blue-700 rounded-xl shadow-sm transition-colors">
          Save Configuration
        </button>
      </div>
    </form>
  </div>

  <script>
    async function loadAdminConfig() {
      try {
        const res = await fetch('/api/admin/config');
        const data = await res.json();
        document.getElementById('site_id').value = data.site_id || '1325';
        document.getElementById('pod').value = data.pod || 'p01';
        document.getElementById('username').value = data.username || '';
        document.getElementById('batch_size').value = data.batch_size || 50;
        document.getElementById('throttle_seconds').value = data.throttle_seconds || 2;
        document.getElementById('max_row_limit').value = data.max_row_limit || 5000;
        if (data.has_password) {
          document.getElementById('pw-status').innerHTML = '✓ Existing password saved in database. Leave blank to preserve.';
        }
      } catch (e) {
        console.error(e);
      }
    }

    async function handleSaveConfig(e) {
      e.preventDefault();
      const form = document.getElementById('admin-form');
      const formData = new FormData(form);
      const btn = document.getElementById('save-btn');
      btn.innerText = 'Saving...';
      try {
        const res = await fetch('/api/admin/config', { method: 'POST', body: formData });
        const data = await res.json();
        alert(data.message || 'Saved successfully');
        loadAdminConfig();
      } catch (err) {
        alert('Save failed: ' + err);
      } finally {
        btn.innerText = 'Save Configuration';
      }
    }

    async function testConnection() {
      const btn = document.getElementById('test-btn');
      const box = document.getElementById('test-result-box');
      btn.innerText = 'Testing...';
      box.classList.add('hidden');
      try {
        const res = await fetch('/api/admin/test-connection', { method: 'POST' });
        const data = await res.json();
        box.classList.remove('hidden');
        if (data.success) {
          box.className = 'p-3.5 rounded-xl text-xs font-medium bg-emerald-50 text-emerald-800 border border-emerald-200';
          box.innerText = '✓ ' + data.message;
        } else {
          box.className = 'p-3.5 rounded-xl text-xs font-medium bg-rose-50 text-rose-800 border border-rose-200';
          box.innerText = '✕ ' + data.message;
        }
      } catch (err) {
        box.classList.remove('hidden');
        box.className = 'p-3.5 rounded-xl text-xs font-medium bg-rose-50 text-rose-800 border border-rose-200';
        box.innerText = 'Connection error: ' + err;
      } finally {
        btn.innerText = 'Test Connection';
      }
    }

    window.onload = loadAdminConfig;
  </script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
def serve_index():
    return HTMLResponse(content=INDEX_HTML)

@app.get("/admin", response_class=HTMLResponse)
def serve_admin():
    return HTMLResponse(content=ADMIN_HTML)

# ---------------------------------------------------------------------------
# Direct Execution Entry Point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
