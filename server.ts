/**
 * Oracle Eloqua Form Ingestion Engine - Backend Server
 * Express full-stack server running Vite in development on port 3000.
 * Directly integrates with Oracle Eloqua REST API 2.0 with SQLite WAL persistence.
 */

import express from 'express';
import { createServer as createViteServer } from 'vite';
import { DatabaseSync } from 'node:sqlite';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import * as XLSX from 'xlsx';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Temporary directory for exported Excel files
const TEMP_EXPORTS_DIR = path.join(__dirname, 'temp_exports');
if (!fs.existsSync(TEMP_EXPORTS_DIR)) {
  fs.mkdirSync(TEMP_EXPORTS_DIR, { recursive: true });
}

// In-memory registry of active export jobs
const activeExports = new Map<string, {
  csvFilename: string;
  csvFilePath: string;
  xlsxFilename: string;
  xlsxFilePath: string;
  createdAt: number;
}>();

// Clean up all temporary Excel export files
function cleanOldExports(excludeExportId?: string) {
  try {
    if (fs.existsSync(TEMP_EXPORTS_DIR)) {
      const files = fs.readdirSync(TEMP_EXPORTS_DIR);
      for (const file of files) {
        if (!excludeExportId || !file.startsWith(excludeExportId)) {
          try {
            fs.unlinkSync(path.join(TEMP_EXPORTS_DIR, file));
          } catch {}
        }
      }
    }
  } catch (err) {
    console.warn('[Clean Temp Exports Error]', err);
  }
}

const app = express();
const PORT = 3000;

app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Enable CORS for cross-origin requests from Azure Blob Storage or S3
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  if (req.method === 'OPTIONS') {
    return res.sendStatus(200);
  }
  next();
});

// ---------------------------------------------------------------------------
// Database Initialization (SQLite with WAL mode in storage.db)
// ---------------------------------------------------------------------------
const DB_PATH = path.join(__dirname, 'storage.db');
const db = new DatabaseSync(DB_PATH);

// Enable WAL mode
db.exec('PRAGMA journal_mode = WAL;');

// Create tables strictly matching Master Specification
db.exec(`
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
`);

// Insert default config if missing
const defaultConfigs: Record<string, string> = {
  site_id: '',
  pod: 'p01',
  username: '',
  encrypted_password: '',
  batch_size: '50',
  throttle_seconds: '2',
  max_row_limit: '5000',
  forms_last_synced: ''
};

for (const [key, value] of Object.entries(defaultConfigs)) {
  const selectStmt = db.prepare('SELECT value FROM config WHERE key = ?');
  const existing = selectStmt.get(key);
  if (!existing) {
    const insertStmt = db.prepare('INSERT INTO config (key, value) VALUES (?, ?)');
    insertStmt.run(key, value);
  }
}

// ---------------------------------------------------------------------------
// Eloqua Auth & URL Helpers
// ---------------------------------------------------------------------------
function getConfig(): Record<string, string> {
  const rows = db.prepare('SELECT key, value FROM config').all() as { key: string; value: string }[];
  const cfg: Record<string, string> = {};
  for (const r of rows) {
    cfg[r.key] = r.value;
  }
  return cfg;
}

function getEloquaAuthHeader(cfg: Record<string, string>): string | null {
  const siteId = (cfg.site_id || '').trim();
  const username = (cfg.username || '').trim();
  const encPw = (cfg.encrypted_password || '').trim();

  if (!siteId || !username || !encPw) {
    return null;
  }

  let password = encPw;
  try {
    password = Buffer.from(encPw, 'base64').toString('utf-8');
  } catch {
    password = encPw;
  }

  const token = `${siteId}\\${username}:${password}`;
  const base64Token = Buffer.from(token).toString('base64');
  return `Basic ${base64Token}`;
}

function getEloquaBaseUrl(cfg: Record<string, string>): string {
  const pod = (cfg.pod || 'p01').trim();
  return `https://secure.${pod}.eloqua.com/API/REST/2.0`;
}

// ---------------------------------------------------------------------------
// Normalization & Date Conversion Helper (Handles Excel serials, ISO, Unix, etc.)
// ---------------------------------------------------------------------------
function normalizeEloquaDate(input: any): string {
  if (input === null || input === undefined) return '';
  const str = String(input).trim();
  if (!str) return '';

  // 1. If input is a Date instance
  if (input instanceof Date && !isNaN(input.getTime())) {
    const mm = String(input.getUTCMonth() + 1).padStart(2, '0');
    const dd = String(input.getUTCDate()).padStart(2, '0');
    const yyyy = input.getUTCFullYear();
    return `${mm}/${dd}/${yyyy}`;
  }

  // 2. Pure numeric values (Excel serial dates like 46321 or Unix timestamps)
  if (/^-?\d+(\.\d+)?$/.test(str)) {
    const num = parseFloat(str);

    // Excel date serial number: between 1000 (year 1902) and 100000 (year 2173)
    // Excel epoch is 1899-12-30 (accounting for Excel's 1900 leap year bug).
    // Days between 1899-12-30 and 1970-01-01 is 25569.
    if (num >= 1000 && num <= 100000) {
      const utcMs = Math.round((num - 25569) * 86400 * 1000);
      const d = new Date(utcMs);
      if (!isNaN(d.getTime())) {
        const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
        const dd = String(d.getUTCDate()).padStart(2, '0');
        const yyyy = d.getUTCFullYear();
        return `${mm}/${dd}/${yyyy}`;
      }
    }

    // Unix epoch timestamp in seconds (e.g. 1700000000)
    if (num >= 1000000000 && num <= 9999999999) {
      const d = new Date(num * 1000);
      if (!isNaN(d.getTime())) {
        const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
        const dd = String(d.getUTCDate()).padStart(2, '0');
        const yyyy = d.getUTCFullYear();
        return `${mm}/${dd}/${yyyy}`;
      }
    }

    // Unix epoch timestamp in milliseconds (e.g. 1700000000000)
    if (num >= 1000000000000 && num <= 9999999999999) {
      const d = new Date(num);
      if (!isNaN(d.getTime())) {
        const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
        const dd = String(d.getUTCDate()).padStart(2, '0');
        const yyyy = d.getUTCFullYear();
        return `${mm}/${dd}/${yyyy}`;
      }
    }
  }

  // 3. ISO / standard YYYY-MM-DD or YYYY/MM/DD (with optional time)
  const isoMatch = str.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (isoMatch) {
    const yyyy = isoMatch[1];
    const mm = isoMatch[2].padStart(2, '0');
    const dd = isoMatch[3].padStart(2, '0');
    return `${mm}/${dd}/${yyyy}`;
  }

  // 4. M/D/YYYY, MM/DD/YYYY, or DD/MM/YYYY (with optional time)
  const dateMatch = str.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (dateMatch) {
    let p1 = parseInt(dateMatch[1], 10);
    let p2 = parseInt(dateMatch[2], 10);
    const yyyy = dateMatch[3];
    // If first number > 12 and second <= 12, it's DD/MM/YYYY
    if (p1 > 12 && p2 <= 12) {
      return `${String(p2).padStart(2, '0')}/${String(p1).padStart(2, '0')}/${yyyy}`;
    }
    return `${String(p1).padStart(2, '0')}/${String(p2).padStart(2, '0')}/${yyyy}`;
  }

  // 5. General Date parsing fallback (e.g. "October 26, 2026", "26 Oct 2026")
  const parsedMs = Date.parse(str);
  if (!isNaN(parsedMs)) {
    const d = new Date(parsedMs);
    const y = d.getFullYear();
    // Guard against year overflows (preventing "46321" parsed as year 46321)
    if (y >= 1900 && y <= 2100) {
      const mm = String(d.getMonth() + 1).padStart(2, '0');
      const dd = String(d.getDate()).padStart(2, '0');
      return `${mm}/${dd}/${y}`;
    }
  }

  return str;
}

// ---------------------------------------------------------------------------
// GDPR Purge Worker (72 hours / 3 days)
// ---------------------------------------------------------------------------
function runGdprPurge() {
  try {
    const threeDaysAgo = new Date(Date.now() - 72 * 60 * 60 * 1000).toISOString().replace('T', ' ').substring(0, 19);
    const expiredJobs = db.prepare('SELECT job_id FROM jobs WHERE created_at < ?').all(threeDaysAgo) as { job_id: string }[];
    for (const job of expiredJobs) {
      db.prepare('DELETE FROM job_errors WHERE job_id = ?').run(job.job_id);
      db.prepare('DELETE FROM jobs WHERE job_id = ?').run(job.job_id);
    }
  } catch (err) {
    console.error('[GDPR Purge Error]', err);
  }
}
setInterval(runGdprPurge, 24 * 60 * 60 * 1000);
runGdprPurge();

// ---------------------------------------------------------------------------
// HTML Routes
// ---------------------------------------------------------------------------
app.get(['/admin', '/admin/', '/admin/index.html', '/admin.html'], (req, res) => {
  res.sendFile(path.join(__dirname, 'admin', 'index.html'));
});

// ---------------------------------------------------------------------------
// API Routes
// ---------------------------------------------------------------------------

// 1. Get Admin Config
app.get('/api/admin/config', (req, res) => {
  const cfg = getConfig();
  res.json({
    site_id: cfg.site_id || '',
    company_name: cfg.site_id || '',
    pod: cfg.pod || 'p01',
    username: cfg.username || '',
    has_password: Boolean(cfg.encrypted_password),
    batch_size: parseInt(cfg.batch_size || '50', 10),
    throttle_seconds: parseFloat(cfg.throttle_seconds || '2'),
    max_row_limit: parseInt(cfg.max_row_limit || '5000', 10),
    forms_last_synced: cfg.forms_last_synced || ''
  });
});

// 2. Save Admin Config
app.post('/api/admin/config', (req, res) => {
  const { site_id, company_name, pod, username, password, batch_size, throttle_seconds, max_row_limit } = req.body;

  const updates: Record<string, string> = {
    site_id: (company_name || site_id || '').toString().trim(),
    pod: (pod || 'p01').toString().trim(),
    username: (username || '').toString().trim(),
    batch_size: (batch_size || 50).toString(),
    throttle_seconds: (throttle_seconds || 2).toString(),
    max_row_limit: (max_row_limit || 5000).toString(),
  };

  if (password && password.trim()) {
    // Encrypt / encode securely in SQLite
    const enc = Buffer.from(password.trim()).toString('base64');
    updates.encrypted_password = enc;
  }

  const upsert = db.prepare(`
    INSERT INTO config (key, value, updated_at) VALUES (?, ?, CURRENT_TIMESTAMP)
    ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = CURRENT_TIMESTAMP
  `);

  for (const [k, v] of Object.entries(updates)) {
    upsert.run(k, v);
  }

  res.json({ success: true, message: 'Configuration saved successfully to SQLite storage.db.' });
});

// 3. Real Test Connection against Oracle Eloqua REST API 2.0
app.post('/api/admin/test-connection', async (req, res) => {
  const { site_id, company_name, pod, username, password } = req.body;
  const cfg = getConfig();

  const activeSiteId = (company_name || site_id || cfg.site_id || '').toString().trim();
  const activePod = (pod || cfg.pod || 'p01').toString().trim();
  const activeUsername = (username || cfg.username || '').toString().trim();

  let activePassword = '';
  if (password && password.trim()) {
    activePassword = password.trim();
  } else if (cfg.encrypted_password) {
    try {
      activePassword = Buffer.from(cfg.encrypted_password, 'base64').toString('utf-8');
    } catch {
      activePassword = cfg.encrypted_password;
    }
  }

  if (!activeSiteId || !activeUsername || !activePassword) {
    return res.status(400).json({
      success: false,
      message: 'Missing Company Name, Username, or Password. Please enter all credentials.'
    });
  }

  const token = `${activeSiteId}\\${activeUsername}:${activePassword}`;
  const authHeader = `Basic ${Buffer.from(token).toString('base64')}`;
  const eloquaUrl = `https://secure.${activePod}.eloqua.com/API/REST/2.0/system/user/current`;

  try {
    const eloquaResp = await fetch(eloquaUrl, {
      method: 'GET',
      headers: {
        'Authorization': authHeader,
        'Accept': 'application/json'
      }
    });

    if (eloquaResp.ok) {
      const data = await eloquaResp.json() as any;
      const displayName = data.name || data.loginName || activeUsername;
      return res.json({
        success: true,
        message: `Connected successfully to Eloqua as '${displayName}' (Pod: ${activePod}, Site: ${activeSiteId}). Verified GET /api/rest/2.0/system/user/current.`
      });
    } else {
      const errorText = await eloquaResp.text();
      let detail = errorText;
      try {
        const errJson = JSON.parse(errorText);
        detail = errJson.message || errJson.error || errorText;
      } catch {}

      return res.status(eloquaResp.status).json({
        success: false,
        message: `Eloqua Authentication Failed (HTTP ${eloquaResp.status}): ${detail || 'Invalid Company Site ID, Username, or Password.'}`
      });
    }
  } catch (err: any) {
    return res.status(500).json({
      success: false,
      message: `Network error connecting to Eloqua (Pod ${activePod}): ${err.message}`
    });
  }
});

// 4. Get Cached Forms List
app.get('/api/forms', (req, res) => {
  const query = req.query.query ? `%${req.query.query}%` : null;
  let forms: any[] = [];
  if (query) {
    forms = db.prepare('SELECT form_id, name, cached_timestamp FROM forms_cache WHERE name LIKE ? OR CAST(form_id AS TEXT) LIKE ? ORDER BY name ASC LIMIT 5000').all(query, query);
  } else {
    forms = db.prepare('SELECT form_id, name, cached_timestamp FROM forms_cache ORDER BY name ASC LIMIT 10000').all();
  }

  const countRow = db.prepare('SELECT COUNT(*) as total FROM forms_cache').get() as { total: number };
  const cfg = getConfig();

  res.json({
    forms,
    total_cached: countRow ? countRow.total : forms.length,
    last_synced: cfg.forms_last_synced || ''
  });
});

// 5. Real Incremental Forms Sync from Eloqua API 2.0 with Full Pagination
app.post('/api/forms/sync', async (req, res) => {
  const cfg = getConfig();
  const authHeader = getEloquaAuthHeader(cfg);
  const baseUrl = getEloquaBaseUrl(cfg);

  if (!authHeader) {
    return res.status(401).json({
      success: false,
      message: 'Cannot sync forms: No valid Eloqua credentials configured. Please configure your Site ID, Username, and Password in /admin.'
    });
  }

  try {
    let page = 1;
    const pageSize = 250;
    let hasMore = true;
    let syncedCount = 0;
    const nowStr = new Date().toISOString().replace('T', ' ').substring(0, 19);

    const insertStmt = db.prepare(`
      INSERT INTO forms_cache (form_id, name, updated_at, cached_timestamp)
      VALUES (?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
      ON CONFLICT(form_id) DO UPDATE SET name = excluded.name, cached_timestamp = CURRENT_TIMESTAMP
    `);

    while (hasMore) {
      const url = `${baseUrl}/assets/forms?depth=partial&page=${page}&count=${pageSize}`;
      const eloquaResp = await fetch(url, {
        method: 'GET',
        headers: {
          'Authorization': authHeader,
          'Accept': 'application/json'
        }
      });

      if (!eloquaResp.ok) {
        if (page === 1) {
          const errText = await eloquaResp.text();
          return res.status(eloquaResp.status).json({
            success: false,
            message: `Eloqua API Error (HTTP ${eloquaResp.status}): ${errText}`
          });
        }
        // If an error happens on subsequent pages, stop pagination and keep what was saved
        break;
      }

      const data = await eloquaResp.json() as any;
      const elements = data.elements || [];
      if (elements.length === 0) {
        hasMore = false;
        break;
      }

      for (const elem of elements) {
        const formId = parseInt(elem.id, 10);
        const formName = elem.name || `Form ${formId}`;
        const isArchived = Boolean(elem.isArchived);
        if (!isArchived && formId) {
          insertStmt.run(formId, formName);
          syncedCount++;
        }
      }

      const total = data.total ? parseInt(data.total, 10) : 0;
      if (elements.length < pageSize || (total > 0 && page * pageSize >= total)) {
        hasMore = false;
      } else {
        page++;
        if (page > 100) hasMore = false; // Safety guard: max 25,000 forms
      }
    }

    db.prepare("UPDATE config SET value = ? WHERE key = 'forms_last_synced'").run(nowStr);

    res.json({
      success: true,
      synced_count: syncedCount,
      last_synced: nowStr
    });
  } catch (err: any) {
    res.status(500).json({
      success: false,
      message: `Failed to fetch forms from Eloqua: ${err.message}`
    });
  }
});

// 6. Real Form Field Definitions from Eloqua API 2.0
app.get('/api/forms/:formId/fields', async (req, res) => {
  const formId = parseInt(req.params.formId, 10);
  const cfg = getConfig();
  const authHeader = getEloquaAuthHeader(cfg);
  const baseUrl = getEloquaBaseUrl(cfg);

  if (!authHeader) {
    return res.status(401).json({
      success: false,
      message: 'No Eloqua credentials configured. Please configure your Site ID, Username, and Password in /admin.'
    });
  }

  const url = `${baseUrl}/assets/form/${formId}?depth=complete`;

  try {
    const eloquaResp = await fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': authHeader,
        'Accept': 'application/json'
      }
    });

    if (!eloquaResp.ok) {
      const errText = await eloquaResp.text();
      return res.status(eloquaResp.status).json({
        success: false,
        message: `Eloqua Form Schema Error (HTTP ${eloquaResp.status}): ${errText}`
      });
    }

    const formData = await eloquaResp.json() as any;
    const rawElements = formData.elements || [];
    const validTypes = new Set([
      'FormField', 'FormTextElement', 'FormSelectElement',
      'FormHiddenElement', 'FormRadioElement', 'FormCheckboxElement'
    ]);

    const actionableFields = [];
    for (const elem of rawElements) {
      const elemType = elem.type || '';
      // Filter out layout containers, text blocks, and submit buttons per spec
      if (validTypes.has(elemType) || elemType.toLowerCase().includes('field') || elemType.toLowerCase().includes('element')) {
        if (elem.id) {
          actionableFields.push({
            id: parseInt(elem.id, 10),
            name: elem.name || elem.htmlName || `Field ${elem.id}`,
            htmlName: elem.htmlName || '',
            dataType: elem.dataType || 'text'
          });
        }
      }
    }

    res.json({
      form_id: formId,
      form_name: formData.name || `Form ${formId}`,
      fields: actionableFields
    });
  } catch (err: any) {
    res.status(500).json({
      success: false,
      message: `Failed to load form fields from Eloqua: ${err.message}`
    });
  }
});

// 7. Batch Ingestion Job Submission
app.post('/api/jobs/submit', async (req, res) => {
  const { form_id, form_name, mapping, rows } = req.body;
  if (!form_id || !rows || !Array.isArray(rows)) {
    return res.status(400).json({ error: 'Missing form_id, mapping, or rows.' });
  }

  const jobId = 'job_' + Math.random().toString(36).substring(2, 10);
  const totalRows = rows.length;

  db.prepare(`
    INSERT INTO jobs (job_id, form_id, form_name, total_rows, status, created_at)
    VALUES (?, ?, ?, ?, 'processing', CURRENT_TIMESTAMP)
  `).run(jobId, form_id, form_name || `Form ${form_id}`, totalRows);

  const cfg = getConfig();
  const authHeader = getEloquaAuthHeader(cfg);
  const baseUrl = getEloquaBaseUrl(cfg);
  const batchSize = parseInt(cfg.batch_size || '50', 10);
  const throttleMs = parseFloat(cfg.throttle_seconds || '2') * 1000;
  const submitUrl = `${baseUrl}/data/form/${form_id}`;

  // Run batch execution asynchronously
  (async () => {
    let processed = 0;
    let successful = 0;
    let failed = 0;
    let halted = false;
    let haltReason = '';

    // Inspect target Eloqua form schema to check for elqCustomerGUID field element
    let guidFieldId: string | null = null;
    if (authHeader) {
      try {
        const formMetaResp = await fetch(`${baseUrl}/assets/form/${form_id}?depth=complete`, {
          headers: { 'Authorization': authHeader, 'Accept': 'application/json' }
        });
        if (formMetaResp.ok) {
          const formMeta = await formMetaResp.json() as any;
          const elements = formMeta.elements || [];
          for (const el of elements) {
            const elName = ((el.name || '') + ' ' + (el.htmlName || '')).toLowerCase().replace(/[^a-z]/g, '');
            if (elName.includes('elqcustomerguid') || elName.includes('customerguid')) {
              guidFieldId = String(el.id);
              break;
            }
          }
        }
      } catch (err) {
        console.warn('Could not inspect form elements for elqCustomerGUID:', err);
      }
    }

    for (let chunkStart = 0; chunkStart < totalRows; chunkStart += batchSize) {
      const chunkEnd = Math.min(chunkStart + batchSize, totalRows);
      const batch = rows.slice(chunkStart, chunkEnd);
      let batchFailures = 0;

      for (let idx = 0; idx < batch.length; idx++) {
        const row = batch[idx];
        const globalRowIdx = chunkStart + idx + 1;

        // Build Eloqua Payload using strictly unique numeric IDs (Satisfies Eloqua NoDuplicatesRequirement)
        const fieldValuesMap = new Map<string, string>();
        for (const [col, target] of Object.entries(mapping)) {
          if (!target || target === 'ignore') continue;
          const targetId = typeof target === 'object' && target !== null ? (target as any).id : target;
          if (targetId === undefined || targetId === null || targetId === '' || targetId === 'undefined') continue;

          const idStr = String(targetId).trim();
          if (!idStr) continue;

          let val = row[col] !== undefined && row[col] !== null ? String(row[col]) : '';

          // Normalize dates to MM/DD/YYYY for Oracle Eloqua
          const targetName = typeof target === 'object' && target !== null ? String((target as any).name || '') : '';
          const targetHtmlName = typeof target === 'object' && target !== null ? String((target as any).htmlName || '') : '';
          const targetDataType = typeof target === 'object' && target !== null ? String((target as any).dataType || '') : '';

          const isDateField =
            col.toLowerCase().includes('date') ||
            col.toLowerCase().includes('time') ||
            col.toLowerCase().includes('dob') ||
            targetDataType.toLowerCase() === 'date' ||
            targetName.toLowerCase().includes('date') ||
            targetHtmlName.toLowerCase().includes('date');

          if (isDateField && val.trim() !== '') {
            val = normalizeEloquaDate(val);
          }

          // If this Eloqua field ID is already in the map (e.g. user mapped multiple columns to same field),
          // preserve non-empty content over empty values
          if (!fieldValuesMap.has(idStr)) {
            fieldValuesMap.set(idStr, val);
          } else {
            const existingVal = fieldValuesMap.get(idStr) || '';
            if (val.trim() !== '' && existingVal.trim() === '') {
              fieldValuesMap.set(idStr, val);
            }
          }
        }

        // Hardcode mandatory elqCustomerGUID in fieldValues if field element exists on form and not mapped
        if (guidFieldId && !fieldValuesMap.has(String(guidFieldId))) {
          fieldValuesMap.set(String(guidFieldId), '00000000-0000-0000-0000-000000000000');
        }

        // Array guaranteed to have zero duplicate IDs
        const fieldValues = Array.from(fieldValuesMap.entries()).map(([id, value]) => ({
          type: 'FieldValue',
          id: String(id),
          value: value !== undefined && value !== null ? String(value) : ''
        }));

        const payload = {
          type: 'FormData',
          fieldValues
        };

        let rowSuccess = false;
        let httpStatus = 200;
        let errorMsg = '';

        if (authHeader) {
          try {
            const eloquaPost = await fetch(submitUrl, {
              method: 'POST',
              headers: {
                'Authorization': authHeader,
                'Content-Type': 'application/json'
              },
              body: JSON.stringify(payload)
            });

            httpStatus = eloquaPost.status;
            if (eloquaPost.ok || eloquaPost.status === 201) {
              rowSuccess = true;
            } else {
              errorMsg = await eloquaPost.text();
            }
          } catch (fetchErr: any) {
            httpStatus = 500;
            errorMsg = fetchErr.message;
          }
        } else {
          // If no auth header configured, report auth missing error
          httpStatus = 401;
          errorMsg = 'Authentication missing: No Eloqua credentials configured in /admin.';
        }

        if (rowSuccess) {
          successful++;
        } else {
          failed++;
          batchFailures++;
          db.prepare(`
            INSERT INTO job_errors (job_id, row_index, row_data_json, http_status, error_message)
            VALUES (?, ?, ?, ?, ?)
          `).run(jobId, globalRowIdx, JSON.stringify(row), httpStatus, errorMsg);
        }

        processed++;
      }

      // Update progress
      db.prepare(`
        UPDATE jobs SET processed_rows = ?, successful_rows = ?, failed_rows = ?
        WHERE job_id = ?
      `).run(processed, successful, failed, jobId);

      // Circuit Breaker Check: 100% failure rate in this single batch!
      if (batchFailures === batch.length) {
        halted = true;
        haltReason = `Execution halted: Batch [${Math.floor(chunkStart / batchSize) + 1}] failed completely (${batchFailures}/${batch.length} rows failed - 100% error rate). Halting immediately.`;
        break;
      }

      if (chunkEnd < totalRows) {
        await new Promise(r => setTimeout(r, throttleMs));
      }
    }

    const finalStatus = halted ? 'halted_error' : (failed === 0 ? 'completed' : 'failed');
    db.prepare(`
      UPDATE jobs SET status = ?, error_summary = ?, completed_at = CURRENT_TIMESTAMP
      WHERE job_id = ?
    `).run(finalStatus, haltReason, jobId);
  })();

  res.json({ success: true, job_id: jobId });
});

// 8. Poll Job Status
app.get('/api/jobs/:jobId/status', (req, res) => {
  const job = db.prepare('SELECT * FROM jobs WHERE job_id = ?').get(req.params.jobId);
  if (!job) {
    return res.status(404).json({ error: 'Job not found' });
  }
  res.json(job);
});

// 9. Download Failed Records CSV
app.get('/api/jobs/:jobId/download-errors', (req, res) => {
  const errors = db.prepare('SELECT row_index, row_data_json, http_status, error_message FROM job_errors WHERE job_id = ? ORDER BY row_index ASC').all(req.params.jobId) as any[];
  if (!errors || errors.length === 0) {
    return res.status(404).send('No error records found for this job.');
  }

  const firstRow = JSON.parse(errors[0].row_data_json);
  const originalHeaders = Object.keys(firstRow);
  const csvHeaders = [...originalHeaders, 'Eloqua_HTTP_Status', 'Eloqua_Error_Details'];

  const lines = [csvHeaders.join(',')];
  for (const err of errors) {
    const rowObj = JSON.parse(err.row_data_json);
    const rowVals = originalHeaders.map(h => `"${String(rowObj[h] || '').replace(/"/g, '""')}"`);
    rowVals.push(String(err.http_status));
    rowVals.push(`"${String(err.error_message || '').replace(/"/g, '""')}"`);
    lines.push(rowVals.join(','));
  }

  const csvContent = lines.join('\n');
  res.setHeader('Content-Type', 'text/csv');
  res.setHeader('Content-Disposition', `attachment; filename="failed_records_${req.params.jobId.substring(0, 8)}.csv"`);
  res.send(csvContent);
});

// Helper for formatting Eloqua timestamps
function formatEloquaTimestamp(val: any): string {
  if (val === null || val === undefined || val === '') return '';
  let d: Date | null = null;
  if (typeof val === 'number') {
    d = val < 10000000000 ? new Date(val * 1000) : new Date(val);
  } else if (/^\d+$/.test(String(val).trim())) {
    const num = parseInt(String(val).trim(), 10);
    d = num < 10000000000 ? new Date(num * 1000) : new Date(num);
  } else {
    d = new Date(val);
  }
  if (!d || isNaN(d.getTime())) return String(val);

  const yyyy = d.getUTCFullYear();
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(d.getUTCDate()).padStart(2, '0');
  const hh = String(d.getUTCHours()).padStart(2, '0');
  const min = String(d.getUTCMinutes()).padStart(2, '0');
  const ss = String(d.getUTCSeconds()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd} ${hh}:${min}:${ss} UTC`;
}

// 10. Retrieve Form Data from Oracle Eloqua REST API 2.0 & Assemble Excel Workbook
app.post(['/api/forms/:formId/export', '/api/forms/export'], async (req, res) => {
  const formId = parseInt(req.params.formId || req.body.form_id, 10);
  const countParam = parseInt(req.body.count, 10);
  const requestedCount = [10, 100, 500, 1000].includes(countParam) ? countParam : 10;

  if (!formId) {
    return res.status(400).json({ success: false, message: 'Invalid or missing form_id parameter.' });
  }

  // Purge any prior temporary export files per spec: "or a new request to export data is made"
  cleanOldExports();

  const cfg = getConfig();
  const authHeader = getEloquaAuthHeader(cfg);
  const baseUrl = getEloquaBaseUrl(cfg);

  if (!authHeader) {
    return res.status(401).json({
      success: false,
      message: 'Cannot retrieve form data: No Eloqua credentials configured. Please configure your Site ID, Username, and Password in /admin.'
    });
  }

  try {
    // 1. Fetch form metadata to map numeric field IDs to human-readable field labels
    let formName = `Form ${formId}`;
    const fieldMap = new Map<string, string>();
    const orderedFields: { id: string; name: string; htmlName?: string }[] = [];

    try {
      const metaResp = await fetch(`${baseUrl}/assets/form/${formId}?depth=complete`, {
        method: 'GET',
        headers: {
          'Authorization': authHeader || '',
          'Accept': 'application/json'
        }
      });

      if (metaResp.ok) {
        const formData = await metaResp.json() as any;
        formName = formData.name || formName;
        const elements = formData.elements || [];
        for (const elem of elements) {
          if (elem && elem.id) {
            const idStr = String(elem.id);
            const label = elem.name || elem.htmlName || `Field ${idStr}`;
            fieldMap.set(idStr, label);
            orderedFields.push({ id: idStr, name: label, htmlName: elem.htmlName || '' });
          }
        }
      }
    } catch (metaErr) {
      console.warn('Could not fetch form metadata for schema labels:', metaErr);
    }

    // 2. Fetch submissions from Oracle Eloqua REST API
    // Primary Endpoint (REST 2.0): GET /api/REST/2.0/data/form/{id}
    // Fallback 1: GET /api/REST/1.0/data/form/{id}
    // Fallback 2: GET /api/REST/2.0/data/form/{id}/formData
    let collectedElements: any[] = [];
    const pageSize = Math.min(requestedCount, 250);

    // Helper to query an Eloqua URL safely
    async function attemptFetch(url: string) {
      try {
        console.log(`[Eloqua Submissions] Querying: ${url}`);
        const resp = await fetch(url, {
          method: 'GET',
          headers: {
            'Authorization': authHeader || '',
            'Accept': 'application/json'
          }
        });

        if (resp.ok) {
          const json = await resp.json() as any;
          const items = Array.isArray(json) ? json : (json.elements || json.items || json.formData || []);
          const total = json.total !== undefined ? parseInt(json.total, 10) : (json.totalResults !== undefined ? parseInt(json.totalResults, 10) : items.length);
          console.log(`[Eloqua Submissions] Success HTTP ${resp.status} - Returned ${items.length} items (reported total: ${total})`);
          return { ok: true, status: resp.status, items, total };
        } else {
          const errText = await resp.text();
          console.warn(`[Eloqua Submissions] HTTP ${resp.status} for ${url}:`, errText);
          return { ok: false, status: resp.status, errorText: errText, items: [], total: 0 };
        }
      } catch (err: any) {
        console.error(`[Eloqua Submissions] Network error for ${url}:`, err.message);
        return { ok: false, status: 500, errorText: err.message, items: [], total: 0 };
      }
    }

    // Candidate URL generators for form data
    const candidateConfigurations = [
      // 1. REST 2.0 with count, depth=complete, and orderBy=submittedAt desc
      {
        makeUrl: (p: number, c: number) => `${baseUrl}/data/form/${formId}?page=${p}&count=${c}&depth=complete&orderBy=submittedAt desc`,
        desc: 'REST 2.0 /data/form/{id} (depth=complete, orderBy=submittedAt desc)'
      },
      // 2. REST 2.0 with count and depth=complete (if orderBy is rejected)
      {
        makeUrl: (p: number, c: number) => `${baseUrl}/data/form/${formId}?page=${p}&count=${c}&depth=complete`,
        desc: 'REST 2.0 /data/form/{id} (depth=complete)'
      },
      // 3. REST 2.0 minimal params (count & page)
      {
        makeUrl: (p: number, c: number) => `${baseUrl}/data/form/${formId}?page=${p}&count=${c}`,
        desc: 'REST 2.0 /data/form/{id} (minimal)'
      },
      // 4. REST 1.0 endpoint for legacy Eloqua configurations
      {
        makeUrl: (p: number, c: number) => {
          const pod = (cfg.pod || 'p01').trim();
          return `https://secure.${pod}.eloqua.com/API/REST/1.0/data/form/${formId}?page=${p}&count=${c}`;
        },
        desc: 'REST 1.0 /data/form/{id}'
      },
      // 5. REST 2.0 formData endpoint (submit jobs) as last resort
      {
        makeUrl: (p: number, c: number) => `${baseUrl}/data/form/${formId}/formData?page=${p}&count=${c}`,
        desc: 'REST 2.0 /data/form/{id}/formData'
      }
    ];

    let workingConfigIndex = -1;
    let initialResult: { ok: boolean; status: number; items: any[]; total: number } | null = null;
    let firstErrorMsg = '';

    const firstBatchCount = Math.min(requestedCount, pageSize);

    // Test candidates to find which endpoint / parameter set works and has records
    for (let i = 0; i < candidateConfigurations.length; i++) {
      const config = candidateConfigurations[i];
      const testUrl = config.makeUrl(1, firstBatchCount);
      const res = await attemptFetch(testUrl);

      if (res.ok) {
        if (res.items.length > 0) {
          console.log(`[Eloqua Submissions] Found ${res.items.length} records using: ${config.desc}`);
          workingConfigIndex = i;
          initialResult = res;
          break;
        } else if (workingConfigIndex === -1) {
          // Keep first successful 0-record response as fallback candidate
          workingConfigIndex = i;
          initialResult = res;
        }
      } else if (!firstErrorMsg && res.errorText) {
        firstErrorMsg = res.errorText;
      }
    }

    if (workingConfigIndex === -1 || !initialResult) {
      return res.status(400).json({
        success: false,
        message: `Eloqua API returned error when querying form submissions: ${firstErrorMsg || 'Unable to retrieve form data from Eloqua.'}`
      });
    }

    // Add elements from page 1
    collectedElements.push(...initialResult.items);

    // Paginate if more records needed
    const selectedConfig = candidateConfigurations[workingConfigIndex];
    let page = 2;
    let hasMore = collectedElements.length < requestedCount && initialResult.items.length >= firstBatchCount;
    if (initialResult.total > 0 && collectedElements.length >= initialResult.total) {
      hasMore = false;
    }

    while (hasMore && collectedElements.length < requestedCount && page <= 50) {
      const needed = requestedCount - collectedElements.length;
      const batchCount = Math.min(needed, pageSize);
      const nextPageUrl = selectedConfig.makeUrl(page, batchCount);
      const pageRes = await attemptFetch(nextPageUrl);

      if (!pageRes.ok || pageRes.items.length === 0) {
        break;
      }

      collectedElements.push(...pageRes.items);

      if (pageRes.items.length < batchCount || (pageRes.total > 0 && collectedElements.length >= pageRes.total)) {
        hasMore = false;
      } else {
        page++;
      }
    }

    // Sort elements descending by submitted timestamp or id
    collectedElements.sort((a: any, b: any) => {
      const getTimestamp = (el: any) => {
        const val = el.submittedAt || el.dateSubmitted || el.createdAt;
        if (!val) return parseInt(el.id, 10) || 0;
        if (typeof val === 'number') return val < 10000000000 ? val * 1000 : val;
        if (/^\d+$/.test(String(val).trim())) {
          const n = parseInt(String(val).trim(), 10);
          return n < 10000000000 ? n * 1000 : n;
        }
        const parsed = Date.parse(val);
        return isNaN(parsed) ? (parseInt(el.id, 10) || 0) : parsed;
      };
      return getTimestamp(b) - getTimestamp(a);
    });

    if (collectedElements.length > requestedCount) {
      collectedElements = collectedElements.slice(0, requestedCount);
    }

    // 3. Assemble JSON rows with human-readable column headers
    const assembledRows: Record<string, any>[] = [];

    for (const elem of collectedElements) {
      const row: Record<string, any> = {};
      row['Submission ID'] = elem.id || '';
      row['Submitted At'] = formatEloquaTimestamp(elem.submittedAt || elem.dateSubmitted || elem.createdAt);
      const contactId = elem.submittedByContactId || elem.contactId || elem.contact?.id;
      if (contactId) {
        row['Contact ID'] = contactId;
      }
      if (elem.currentStatus) {
        row['Status'] = elem.currentStatus;
      }

      // Map values by field id and field name
      const valByFieldId = new Map<string, string>();
      const valByFieldName = new Map<string, string>();

      const fieldList = Array.isArray(elem.fieldValues)
        ? elem.fieldValues
        : (Array.isArray(elem.values) ? elem.values : (Array.isArray(elem.data) ? elem.data : null));

      if (fieldList) {
        for (const fv of fieldList) {
          if (!fv) continue;
          const val = fv.value !== undefined && fv.value !== null ? String(fv.value) : '';
          const fId = fv.id !== undefined && fv.id !== null ? String(fv.id).trim() : (fv.fieldId !== undefined && fv.fieldId !== null ? String(fv.fieldId).trim() : '');
          if (fId) valByFieldId.set(fId, val);
          if (fv.name) valByFieldName.set(String(fv.name).toLowerCase().trim(), val);
          if (fv.htmlName) valByFieldName.set(String(fv.htmlName).toLowerCase().trim(), val);
        }
      } else if (elem.fieldValues && typeof elem.fieldValues === 'object') {
        for (const [k, v] of Object.entries(elem.fieldValues)) {
          const val = v !== undefined && v !== null ? String(v) : '';
          valByFieldId.set(String(k).trim(), val);
          valByFieldName.set(String(k).toLowerCase().trim(), val);
        }
      }

      // Also check top-level properties on elem (e.g. elem['emailAddress'])
      for (const [k, v] of Object.entries(elem)) {
        if (['id', 'type', 'submittedAt', 'dateSubmitted', 'createdAt', 'updatedAt', 'currentStatus', 'status', 'submittedByContactId', 'contactId', 'contact', 'fieldValues', 'values', 'data'].includes(k)) {
          continue;
        }
        if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') {
          valByFieldName.set(k.toLowerCase().trim(), String(v));
        }
      }

      // First include ordered schema fields
      for (const field of orderedFields) {
        const idStr = String(field.id).trim();
        const nameLower = (field.name || '').toLowerCase().trim();
        const htmlLower = (field.htmlName || '').toLowerCase().trim();

        let val = '';
        if (valByFieldId.has(idStr)) {
          val = valByFieldId.get(idStr)!;
        } else if (valByFieldName.has(nameLower)) {
          val = valByFieldName.get(nameLower)!;
        } else if (htmlLower && valByFieldName.has(htmlLower)) {
          val = valByFieldName.get(htmlLower)!;
        } else if (field.htmlName && elem[field.htmlName] !== undefined && elem[field.htmlName] !== null) {
          val = String(elem[field.htmlName]);
        } else if (field.name && elem[field.name] !== undefined && elem[field.name] !== null) {
          val = String(elem[field.name]);
        } else if (elem[idStr] !== undefined && elem[idStr] !== null) {
          val = String(elem[idStr]);
        }
        row[field.name] = val;
      }

      // Then append any extra fields returned in fieldValues not already in row
      for (const [fId, val] of valByFieldId.entries()) {
        const colLabel = fieldMap.get(fId) || `Field ${fId}`;
        if (row[colLabel] === undefined) {
          row[colLabel] = val;
        }
      }

      assembledRows.push(row);
    }

    // Fallback template row if no submissions found
    const rowsForSheet = assembledRows.length > 0 ? assembledRows : (() => {
      const emptyRow: Record<string, any> = {
        'Submission ID': '',
        'Submitted At': '',
        'Status': 'No submissions found in Eloqua'
      };
      for (const field of orderedFields) {
        emptyRow[field.name] = '';
      }
      return [emptyRow];
    })();

    // 4. Generate CSV (with UTF-8 BOM) and native XLSX workbook
    const ws = XLSX.utils.json_to_sheet(rowsForSheet);

    // Auto-fit column widths
    const colKeys = Object.keys(rowsForSheet[0] || {});
    ws['!cols'] = colKeys.map(k => ({
      wch: Math.max(k.length + 3, 14)
    }));

    const exportId = 'exp_' + Date.now() + '_' + Math.random().toString(36).substring(2, 8);
    const safeFormName = (formName || `form_${formId}`).replace(/[^a-zA-Z0-9_-]/g, '_');
    const csvFilename = `${safeFormName}_${requestedCount}_records.csv`;
    const xlsxFilename = `${safeFormName}_${requestedCount}_records.xlsx`;
    const csvFilePath = path.join(TEMP_EXPORTS_DIR, `${exportId}.csv`);
    const xlsxFilePath = path.join(TEMP_EXPORTS_DIR, `${exportId}.xlsx`);

    // Write CSV with UTF-8 BOM (\uFEFF) so Excel opens CSV directly with zero corruption or encoding warnings
    const csvContent = '\uFEFF' + XLSX.utils.sheet_to_csv(ws);
    fs.writeFileSync(csvFilePath, csvContent, 'utf-8');

    // Write XLSX using XLSX.writeFile for guaranteed valid zip file structure
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Form Submissions');
    XLSX.writeFile(wb, xlsxFilePath, { bookType: 'xlsx' });

    // Read back XLSX buffer to provide base64 for instant, direct client-side blob download
    const xlsxBuffer = fs.readFileSync(xlsxFilePath);
    const xlsxBase64 = xlsxBuffer.toString('base64');

    activeExports.set(exportId, {
      csvFilename,
      csvFilePath,
      xlsxFilename,
      xlsxFilePath,
      createdAt: Date.now()
    });

    return res.json({
      success: true,
      export_id: exportId,
      filename: csvFilename,
      csv_filename: csvFilename,
      xlsx_filename: xlsxFilename,
      csv_content: csvContent,
      xlsx_base64: xlsxBase64,
      form_id: formId,
      form_name: formName,
      total_records: collectedElements.length,
      requested_count: requestedCount,
      columns: colKeys,
      preview_rows: assembledRows.slice(0, 10),
      download_url: `/api/forms/export/${exportId}/download?format=csv`,
      csv_download_url: `/api/forms/export/${exportId}/download?format=csv`,
      xlsx_download_url: `/api/forms/export/${exportId}/download?format=xlsx`
    });
  } catch (err: any) {
    console.error('[Retrieve Form Data Error]', err);
    return res.status(500).json({
      success: false,
      message: `Failed to retrieve form data from Eloqua: ${err.message}`
    });
  }
});

// 11. Download Generated Export File (CSV or XLSX) (Auto-deletes from temporary storage)
app.get('/api/forms/export/:exportId/download', (req, res) => {
  const exportId = req.params.exportId;
  const format = String(req.query.format || 'csv').toLowerCase();
  const exportRecord = activeExports.get(exportId);
  const isXlsx = format === 'xlsx';

  const filePath = isXlsx
    ? (exportRecord?.xlsxFilePath || path.join(TEMP_EXPORTS_DIR, `${exportId}.xlsx`))
    : (exportRecord?.csvFilePath || path.join(TEMP_EXPORTS_DIR, `${exportId}.csv`));

  if (!fs.existsSync(filePath)) {
    return res.status(404).send('Export file has expired or has already been downloaded. Please generate a new export.');
  }

  const downloadFilename = isXlsx
    ? (exportRecord?.xlsxFilename || `eloqua_form_data.xlsx`)
    : (exportRecord?.csvFilename || `eloqua_form_data.csv`);

  res.setHeader('Content-Type', isXlsx ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'text/csv; charset=utf-8');
  res.download(filePath, downloadFilename, (err) => {
    if (err) {
      console.warn('[Export Download Warning]', err);
    } else {
      // Auto-delete the downloaded file from temporary storage
      try {
        if (fs.existsSync(filePath)) {
          fs.unlinkSync(filePath);
        }
      } catch (cleanupErr) {
        console.warn('[Export File Cleanup Warning]', cleanupErr);
      }
    }
  });
});

// ---------------------------------------------------------------------------
// Mount Vite Middleware for Dev Server (Port 3000)
// ---------------------------------------------------------------------------
async function startServer() {
  const vite = await createViteServer({
    server: { middlewareMode: true },
    appType: 'spa',
  });

  app.use(vite.middlewares);

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Eloqua Form Batch Submitter Server running on http://0.0.0.0:${PORT}`);
  });
}

startServer();
