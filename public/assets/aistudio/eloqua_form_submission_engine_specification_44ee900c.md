# Oracle Eloqua Form Submission Engine - Master Specification

## 1. Project Overview & Scope
The application is a standalone, self-contained web tool built with Python (FastAPI) serving an integrated modern frontend (Tailwind CSS + Vanilla JavaScript). It enables marketing operations teams to upload Excel (`.xlsx`) or CSV files, map column headers to Oracle Eloqua form fields using intelligent fuzzy matching, and submit records in controlled batches directly to Eloqua forms via the Eloqua REST API 2.0.

### Core Guarantees:
- **Zero Plaintext Credentials in UI:** Passwords entered in the admin interface are stored securely in a local SQLite database (base64/reversible encryption for Basic Auth transmission) and are **never** returned or displayed in the UI under any circumstance.
- **Circuit Breaker:** The runner processes rows in batches of 50. If 100% of rows fail within any batch, execution halts immediately to prevent account lockouts or continuous invalid submissions.
- **GDPR Compliance:** Uploaded files, generated error reports, and job records older than 3 days (72 hours) are automatically purged from disk and database.
- **Multi-Tenant / Multi-User Safe:** Every upload session is isolated with a unique UUID (`job_id`), preventing race conditions or data collisions when multiple users access the app concurrently.

---

## 2. Technical Architecture & Stack

### Backend
- **Framework:** Python 3.10+ with `FastAPI` and `uvicorn`.
- **Database:** SQLite with Write-Ahead Logging (`PRAGMA journal_mode=WAL;`) enabled for high-concurrency file access stored in `storage.db`.
- **HTTP Client:** `httpx` or `requests` with connection pooling and timeouts for Eloqua REST API calls.
- **File Parsing:** `openpyxl` (for Excel `.xlsx` first tab) and Python built-in `csv` module.
- **Background Worker:** FastAPI `BackgroundTasks` for asynchronous batch processing and periodic GDPR data purges.

### Frontend
- **Interface:** Single-Page Application (SPA) served directly by FastAPI as a static bundle (`index.html`) matching the provided corporate design system.
- **Styling:** Tailwind CSS (loaded via CDN) adhering to the light-gray/blue card layout, step indicators, top navigation bar, and sticky bottom action bar.
- **Fuzzy Matching:** Client-side string similarity (Levenshtein distance / token sorting) with instant alias matching.

---

## 3. Eloqua API & Connectivity Specification

### Base Credentials & Endpoints
- **Pod:** `p01`
- **Site ID (Company):** `1325`
- **Base REST URL:** `https://secure.p01.eloqua.com/API/REST/2.0`
- **Authentication Scheme:** HTTP Basic Auth
  - Header: `Authorization: Basic <base64(Company\Username:Password)>`
  - Example: `base64("1325\\OpsUser:SecretPass123")`

### Key Endpoints Used

#### 1. Test Connection
- `GET /api/rest/2.0/system/user/current`
- Verifies credentials, returns 200 OK on success.

#### 2. Get Forms List (Cached & Incremental)
- `GET /api/rest/2.0/assets/forms?depth=partial&count=1000&page={page}`
- Filters for active forms only (`isArchived == false` or `type == "Form"`).
- Incremental sync checks for form IDs greater than the maximum `id` currently stored in the local SQLite cache.

#### 3. Get Form Field Definitions
- `GET /api/rest/2.0/assets/form/{formId}?depth=complete`
- **Filtering Rule:** Filter out layout containers, field groups, instructional text blocks, and submit buttons. Only retain actionable input fields:
  - Valid field types: `FormField`, `FormTextElement`, `FormSelectElement`, `FormHiddenElement`, `FormRadioElement`, `FormCheckboxElement`.
  - Extract for each field:
    - `id`: Internal numeric Eloqua Form Field ID (e.g., `14820`). **Mandatory for submission payloads.**
    - `name`: Human-readable display label (e.g., `"First Name"`).
    - `htmlName`: Internal HTML name (e.g., `"firstName1"`).
    - `dataType`: e.g., `text`, `numeric`, `date`.

#### 4. Form Data Submission
- `POST /api/rest/2.0/data/form/{formId}`
- **Content-Type:** `application/json`
- **Payload Schema:**
```json
{
  "type": "FormData",
  "fieldValues": [
    {
      "type": "FieldValue",
      "id": "14820",
      "value": "jane.doe@example.com"
    },
    {
      "type": "FieldValue",
      "id": "14821",
      "value": "Jane"
    }
  ]
}
```
- **Blank Values:** Mapped fields containing empty spreadsheet cells must still be passed with an empty string `""` to preserve field mapping structure.

---

## 4. Database Schema (`storage.db`)

```sql
PRAGMA journal_mode = WAL;

-- App and Eloqua Configuration
CREATE TABLE IF NOT EXISTS config (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

-- Eloqua Cached Forms
CREATE TABLE IF NOT EXISTS forms_cache (
    form_id INTEGER PRIMARY KEY,
    name TEXT NOT NULL,
    created_at DATETIME,
    updated_at DATETIME,
    cached_timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
);

-- User Upload Jobs
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

-- Row-Level Error Logs for Failed Submissions
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

### Initial Config Defaults
- `site_id`: `1325`
- `pod`: `p01`
- `username`: `""`
- `encrypted_password`: `""`
- `batch_size`: `50`
- `throttle_seconds`: `2`
- `max_row_limit`: `5000`
- `forms_last_synced`: `NULL`

---

## 5. UI Layout & Visual Design System

The visual theme mirrors the provided clean enterprise design:
- **Background:** Soft grayish blue (`#F3F5F9`).
- **Cards:** Clean white (`#FFFFFF`) with subtle border (`border-gray-200`) and soft rounded corners (`rounded-2xl`).
- **Typography:** Modern sans-serif (Inter / system UI) with strong hierarchical contrast.
- **Top Navigation Bar:**
  - Left: Logo icon + App Title `"Eloqua Form Ingestion Engine"`.
  - Right: Discrete `"Start Over"` icon button.
  - *Note:* The `/admin` path has **no button** in the UI and is accessible only by navigating directly to `/admin`.
- **Footer/Bottom Bar:** Sticky bottom action bar containing contextual status text and primary action button (e.g., `"Continue →"` or `"Submit to Eloqua"`).

### 5-Step Wizard Flow

#### Step 1: Select Form
- Searchable typeahead input allowing filtering across 1,500+ cached forms by name or numeric ID.
- Displays metadata: Last synced date and total cached form count.
- **"Sync Forms" Button:** Triggers incremental fetch of newer forms from Eloqua.
- **"Refresh Fields" Button:** Forces an immediate re-fetch of the selected form's fields to capture any recent schema updates.
- Selection unlocks the Next button.

#### Step 2: Upload File
- Large dashed drop-zone card matching the mockup:
  - Cloud upload icon.
  - `"Drag & drop your file here"` heading with subtext: `"or click anywhere to browse for a .xlsx or .csv file"`.
  - File format badge: `.xlsx or .csv`.
- Constraints: Reads **only the first sheet** of `.xlsx`. Max 5,000 rows. Rejects empty files.

#### Step 3: Map Fields
- **Duplicate Header Check:** Highlights duplicate headers in the sheet (e.g., two columns named `"Notes"`) with yellow warning badges, prompting the user to verify which column is being mapped.
- **Email Validation:** Requires at least one field containing `"email"` (case-insensitive) to be mapped before allowing progression.
- **Fuzzy Auto-Matcher:** Pre-selects Eloqua form fields based on spreadsheet column names using:
  - Exact match (case-insensitive, normalized spaces).
  - Common alias match (e.g., `first` -> `First Name`, `email address` -> `Email Address`, `cell` -> `Mobile Phone`).
  - Levenshtein ratio > 0.70.
- **Data Preview:** Displays sample values from row 1 of the uploaded spreadsheet next to each column header so users can verify alignment.
- **Unmapped Columns:** Clear `[ -- Ignore Column -- ]` option for headers not intended for submission.

#### Step 4: Review Mapping
- Displays a clean two-column summary card showing only active mappings:
  - Column 1: `Uploaded Column Header (with Sample Value)`
  - Column 2: `Eloqua Target Field Name [Field ID]`
- Summary box at bottom: `X columns mapped, Y columns will be ignored. Total rows to submit: Z`.
- Requires explicit user confirmation before initiating the batch process.

#### Step 5: Process & Results
- Dynamic batch progress bar updating every 50 records.
- Status cards: `Total Rows`, `Processed`, `Successful (Green)`, `Failed (Red)`.
- **Completion States:**
  - **100% Success:** Congratulatory banner displaying total processed count; no download button needed.
  - **Partial Errors:** Summary of errors + prominent **"Download Failed Records (.csv)"** button.
  - **Halted via Circuit Breaker:** Prominent red warning banner: `"Execution halted: Batch [X] failed completely (100% error rate). Check credentials, permissions, or required form rules."` + Download button for rows processed up to failure.

---

## 6. Data Normalization & Fuzzy Matching Logic

### 1. Fuzzy Matching Rules
- Normalize text: lowercase, remove special characters (`_`, `-`, `.`), trim extra whitespace.
- Pre-defined dictionary matches:
  - `email`, `e-mail`, `email address`, `emailaddress` ➔ Form field containing `"email"`.
  - `first`, `fname`, `first name`, `firstname` ➔ Form field containing `"first"`.
  - `last`, `lname`, `last name`, `lastname` ➔ Form field containing `"last"`.
  - `phone`, `telephone`, `mobile`, `cell` ➔ Form field containing `"phone"`.
  - `company`, `company name`, `organization` ➔ Form field containing `"company"`.
- If no alias matches, calculate Levenshtein distance against Eloqua field `name`. If score >= 0.75, automatically select; otherwise leave as `[ -- Ignore Column -- ]`.

### 2. Date Formatting
- Any column detected as containing date strings or Excel date serials must be formatted to standard Eloqua format: `MM/DD/YYYY`.
- Implementation: Attempt parsing via Python `dateutil` or standard regex; if valid date, convert to `MM/DD/YYYY`. If parsing fails, pass original string.

---

## 7. Batch Processing & Circuit Breaker Engine

### Batch Parameters
- **Batch Size:** 50 records per chunk.
- **Throttle Delay:** Configurable pause (default: 2 seconds) between batches to respect Eloqua API rate limits.
- **Execution Model:**
  1. The backend reads rows 1 to 50 for the given `job_id`.
  2. Constructs individual JSON payloads for each row using the numerical `id` of each mapped field.
  3. Dispatches row requests asynchronously or via a tight connection pool to `POST /api/rest/2.0/data/form/{formId}`.
  4. Collects HTTP response statuses and error payloads.
  5. **Circuit Breaker Check:**
     $$\text{Batch Failure Rate} = \frac{\text{Failed Rows in Batch}}{50} \times 100\%$$
     If $\text{Batch Failure Rate} == 100\%$:
     - Set job status to `halted_error`.
     - Record system error message.
     - Immediately terminate processing loop. Do not execute subsequent batches.
  6. If at least 1 record succeeds, wait `throttle_seconds`, then proceed to next batch of 50.

---

## 8. Error Reporting Export Specification

When a job completes with failures or halts via circuit breaker, the app allows generating an Error CSV file containing **only the rows that failed**.

### Structure of the Exported CSV
- Columns 1 through $N$: All original columns from the uploaded spreadsheet for that failed row.
- Column $N+1$: `Eloqua_HTTP_Status` (e.g., `400`, `422`, `500`).
- Column $N+2$: `Eloqua_Error_Details` (e.g., `Field validation failed: EmailAddress is not a valid email address format`).

---

## 9. GDPR Retention & Auto-Purge Worker

To comply with data privacy policies and GDPR:
- **Retention Limit:** 72 hours (3 days).
- **Purge Logic:**
  - A background task executes on application startup and every 24 hours.
  - Queries `jobs` table for records where `created_at < datetime('now', '-3 days')`.
  - For each expired job:
    1. Unlinks and deletes the uploaded spreadsheet file from the `uploads/` directory.
    2. Unlinks and deletes any generated error CSV from the `reports/` directory.
    3. Deletes associated entries from `job_errors` and `jobs` tables in `storage.db`.
- Operates silently without interfering with active jobs.

---

## 10. Hidden Admin Interface (`/admin`)

- **Route:** Accessible directly at `/admin` (no visible buttons or links in the main application).
- **Features:**
  - Configuration Form:
    - Eloqua Site ID (default: `1325`)
    - Eloqua Pod (default: `p01`)
    - Username
    - Password (masked input field: `<input type="password">`). Never populated with existing password on load; leaving blank on save preserves current password.
    - Batch Size (number, default: `50`)
    - Batch Throttle in Seconds (number, default: `2`)
    - Max Upload Row Ceiling (number, default: `5000`)
  - **"Test Connection" Action:** Sends a live test request to `GET /api/rest/2.0/system/user/current` using entered credentials and displays an instant success/failure alert.
  - **"Save Configuration" Action:** Commits values to SQLite `config` table. Passwords are saved with two-way encryption or base64 encoding and never printed to response bodies.