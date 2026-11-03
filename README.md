# Eloqua FormOps

A full-stack, enterprise application built with Node.js/Express and TypeScript that integrates directly with **Oracle Eloqua REST API 2.0**. It provides two primary marketing operations workflows:
1. **Submit form data to Eloqua**: Ingest spreadsheets (`.xlsx` or `.csv`), map column headers to Eloqua form fields using fuzzy matching (Levenshtein + aliases), and dispatch records in controlled batches with real-time feedback, circuit breaker safety, and automated GDPR compliance.
2. **Retrieve form data from Eloqua**: Extract recent form submission records via `GET /api/REST/2.0/data/form/{id}/formData`, map numeric field IDs to human-readable form field names, assemble the JSON into a clean Excel (`.xlsx`) workbook, and download with automated temporary storage cleanup.

---

## Key Features

1. **Integrated Single Executable Architecture:**
   - Single-file FastAPI application (`main.py`) serving both API routes and embedded HTML5 / Tailwind CSS frontend.
   - Requires zero Node.js or frontend build steps when running in Python.
   - Built-in SQLite database with Write-Ahead Logging (`PRAGMA journal_mode = WAL;`) for high concurrency.

2. **Full Oracle Eloqua REST API 2.0 Integration:**
   - **Pod:** `p01` (configurable)
   - **Site ID (Company):** `1325` (configurable)
   - **Auth:** HTTP Basic Auth `Basic base64("1325\\<username>:<password>")`.
   - **Field Numeric ID Mapping:** Ingests fields using the required numerical `id` format:
     ```json
     {
       "type": "FormData",
       "fieldValues": [
         { "type": "FieldValue", "id": "14820", "value": "jane.doe@example.com" },
         { "type": "FieldValue", "id": "14821", "value": "Jane" }
       ]
     }
     ```
   - **Date Normalization:** Dates automatically formatted to `MM/DD/YYYY`.
   - **Empty Value Handling:** Empty cells passed as `""` to preserve field structure.

3. **Intelligent 5-Step Ingestion Wizard:**
   - **Step 1 - Select Form:** Searchable typeahead over cached forms. "Sync Forms" (incremental fetch) and "Refresh Fields" (schema re-fetch).
   - **Step 2 - Upload File:** Drag-and-drop `.xlsx` or `.csv`. Reads first sheet only, enforces 5,000 max row ceiling.
   - **Step 3 - Map Fields:** Header fuzzy auto-matching (Levenshtein ratio >= 0.70 + marketing aliases like `fname` -> `First Name`). Highlights duplicate headers with warning badges. Enforces mandatory email mapping.
   - **Step 4 - Review Mapping:** 2-column active mapping table (`Uploaded Column [Sample Value]` -> `Eloqua Field Name [ID]`).
   - **Step 5 - Process & Results:** Dynamic progress bar updated per batch of 50. Circuit breaker halts if 100% of rows fail in any single batch. Download Failed Records (`.csv`) with `Eloqua_HTTP_Status` and `Eloqua_Error_Details`.

4. **Hidden Admin Portal (`/admin`):**
   - Accessible strictly at `/admin` (no navigation link in the user UI).
   - Configure Site ID, Pod, Username, Password, Batch Size (default 50), Throttle Seconds (default 2s), and Max Rows (default 5000).
   - "Test Connection" button calling `GET /api/rest/2.0/system/user/current`.
   - Passwords stored securely and **never** returned to the UI.

5. **GDPR Auto-Purge Worker:**
   - Automated background task purges uploaded spreadsheets, error CSV reports, and database job logs older than 72 hours (3 days).

6. **Multi-User Isolation:**
   - Generates unique UUID (`job_id`) for each upload session, preventing race conditions.

---

## Quickstart & Local Execution

### Prerequisites
- Python 3.10 or higher

### Installation
```bash
# 1. Clone repository & change directory
cd /path/to/project

# 2. Create and activate a virtual environment
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt
```

### Run Server
```bash
uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

Open your browser to:
- **Main Wizard UI:** `http://localhost:8000/`
- **Hidden Admin Panel:** `http://localhost:8000/admin`
- **Interactive Swagger Docs:** `http://localhost:8000/docs`

---

## Database Schema (`storage.db`)

```sql
PRAGMA journal_mode = WAL;

CREATE TABLE IF NOT EXISTS config (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS forms_cache (
    form_id INTEGER PRIMARY KEY,
    name TEXT NOT NULL,
    created_at DATETIME,
    updated_at DATETIME,
    cached_timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
);

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
```

---

## Circuit Breaker Logic

The dispatcher divides records into chunks of 50.
For each chunk:
- If 100% of rows fail in that chunk:
  - Job status is set to `halted_error`.
  - Ingestion terminates immediately.
  - User receives a prominent warning banner and an option to download the error CSV of processed rows.
- If at least 1 record succeeds:
  - Waits `throttle_seconds` (default 2s) and continues to the next batch.
