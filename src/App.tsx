/**
 * Oracle Eloqua Form Ingestion Engine
 * Master Enterprise Application integrating directly with Oracle Eloqua REST API 2.0
 */

import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import {
  FileSpreadsheet,
  CheckCircle2,
  AlertTriangle,
  RefreshCw,
  UploadCloud,
  ChevronRight,
  ArrowLeft,
  RotateCcw,
  Download,
  Search,
  Check,
  AlertCircle,
  Sliders,
  ShieldCheck,
  Lock
} from 'lucide-react';

// --- Interfaces ---
interface EloquaForm {
  form_id: number;
  name: string;
  cached_timestamp?: string;
}

interface EloquaField {
  id: number;
  name: string;
  htmlName: string;
  dataType: string;
}

interface UploadedData {
  job_id: string;
  filename: string;
  headers: string[];
  duplicate_headers: string[];
  total_rows: number;
  sample_row: Record<string, any>;
  rows: Record<string, any>[];
}

interface FailedRow {
  rowIndex: number;
  data: Record<string, any>;
  httpStatus: number;
  errorDetail: string;
}

interface ConfigSettings {
  site_id: string;
  pod: string;
  username: string;
  has_password: boolean;
  batch_size: number;
  throttle_seconds: number;
  max_row_limit: number;
  forms_last_synced: string;
}

// Marketing operations alias mapping for fuzzy matching
const ALIAS_MAP: Record<string, string> = {
  first: "first",
  fname: "first",
  "first name": "first",
  firstname: "first",
  last: "last",
  lname: "last",
  "last name": "last",
  lastname: "last",
  email: "email",
  "email address": "email",
  emailaddress: "email",
  "e-mail": "email",
  phone: "phone",
  telephone: "phone",
  mobile: "phone",
  cell: "phone",
  "bus phone": "phone",
  "business phone": "phone",
  company: "company",
  "company name": "company",
  organization: "company",
  title: "title",
  "job title": "title",
  country: "country",
  comments: "comments",
  notes: "comments",
  date: "date",
  registration: "date",
  "reg date": "date"
};

function levenshteinSimilarity(s1: string, s2: string): number {
  s1 = (s1 || "").toLowerCase().trim();
  s2 = (s2 || "").toLowerCase().trim();
  if (s1 === s2) return 1.0;
  const len1 = s1.length, len2 = s2.length;
  if (!len1 || !len2) return 0.0;
  const dp: number[][] = Array.from({ length: len1 + 1 }, () => Array(len2 + 1).fill(0));
  for (let i = 0; i <= len1; i++) dp[i][0] = i;
  for (let j = 0; j <= len2; j++) dp[0][j] = j;
  for (let i = 1; i <= len1; i++) {
    for (let j = 1; j <= len2; j++) {
      const cost = s1[i - 1] === s2[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1,
        dp[i][j - 1] + 1,
        dp[i - 1][j - 1] + cost
      );
    }
  }
  return 1.0 - dp[len1][len2] / Math.max(len1, len2);
}

/**
 * Normalizes dates to MM/DD/YYYY format for Oracle Eloqua.
 * Handles Excel serial numbers (e.g. 46321 -> 10/26/2026),
 * Unix timestamps (seconds & milliseconds),
 * standard date strings (ISO YYYY-MM-DD, MM/DD/YYYY, DD/MM/YYYY, etc.),
 * and JS Date objects without timezone shift.
 */
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

    // Excel date serial number: roughly between 1000 (year 1902) and 100000 (year 2173)
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

export default function App() {
  // Wizard Navigation (Steps 1 to 5)
  const [currentStep, setCurrentStep] = useState<number>(1);
  const [isAdminView, setIsAdminView] = useState<boolean>(false);

  // Configuration from SQLite
  const [config, setConfig] = useState<ConfigSettings>({
    site_id: "",
    pod: "p01",
    username: "",
    has_password: false,
    batch_size: 50,
    throttle_seconds: 2,
    max_row_limit: 5000,
    forms_last_synced: ""
  });

  // Admin form state
  const [adminSiteId, setAdminSiteId] = useState("");
  const [adminPod, setAdminPod] = useState("p01");
  const [adminUsername, setAdminUsername] = useState("");
  const [adminPassword, setAdminPassword] = useState("");
  const [adminBatchSize, setAdminBatchSize] = useState(50);
  const [adminThrottle, setAdminThrottle] = useState(2);
  const [adminMaxRows, setAdminMaxRows] = useState(5000);
  const [adminTestResult, setAdminTestResult] = useState<{ success: boolean; message: string } | null>(null);
  const [adminSaving, setAdminSaving] = useState(false);
  const [adminTesting, setAdminTesting] = useState(false);

  // Step 1: Forms
  const [formsList, setFormsList] = useState<EloquaForm[]>([]);
  const [totalCachedForms, setTotalCachedForms] = useState<number>(0);
  const [formSearch, setFormSearch] = useState<string>("");
  const [selectedForm, setSelectedForm] = useState<EloquaForm | null>(null);
  const [formFields, setFormFields] = useState<EloquaField[]>([]);
  const [isSyncingForms, setIsSyncingForms] = useState<boolean>(false);
  const [isRefreshingFields, setIsRefreshingFields] = useState<boolean>(false);
  const [formDropdownOpen, setFormDropdownOpen] = useState<boolean>(false);
  const [syncFeedback, setSyncFeedback] = useState<string | null>(null);

  // Step 2: Uploaded file
  const [uploadedData, setUploadedData] = useState<UploadedData | null>(null);
  const [isUploading, setIsUploading] = useState<boolean>(false);
  const [uploadError, setUploadError] = useState<string | null>(null);

  // Step 3: Field Mapping
  const [mapping, setMapping] = useState<Record<string, EloquaField | "ignore">>({});

  // Mode selection: 'select' | 'submit' | 'retrieve'
  const [appMode, setAppMode] = useState<'select' | 'submit' | 'retrieve'>('select');

  // Retrieve form data state
  const [retrieveSelectedForm, setRetrieveSelectedForm] = useState<EloquaForm | null>(null);
  const [retrieveFormSearch, setRetrieveFormSearch] = useState<string>('');
  const [retrieveDropdownOpen, setRetrieveDropdownOpen] = useState<boolean>(false);
  const [retrieveRecordCount, setRetrieveRecordCount] = useState<number>(100);
  const [isRetrieving, setIsRetrieving] = useState<boolean>(false);
  const [retrieveError, setRetrieveError] = useState<string | null>(null);
  const [retrieveResult, setRetrieveResult] = useState<{
    export_id: string;
    filename: string;
    csv_filename?: string;
    xlsx_filename?: string;
    csv_content?: string;
    xlsx_base64?: string;
    total_records: number;
    requested_count: number;
    form_id: number;
    form_name: string;
    columns: string[];
    preview_rows: Record<string, any>[];
    download_url: string;
    csv_download_url?: string;
    xlsx_download_url?: string;
  } | null>(null);

  // Step 5: Processing & Circuit Breaker
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [processedRows, setProcessedRows] = useState<number>(0);
  const [successfulRows, setSuccessfulRows] = useState<number>(0);
  const [failedRows, setFailedRows] = useState<number>(0);
  const [circuitBreakerTriggered, setCircuitBreakerTriggered] = useState<boolean>(false);
  const [circuitBreakerReason, setCircuitBreakerReason] = useState<string>("");
  const [activeJobId, setActiveJobId] = useState<string | null>(null);
  const pollingRef = useRef<NodeJS.Timeout | null>(null);

  // Load configuration and cached forms on startup
  useEffect(() => {
    loadConfig();
    loadCachedForms();

    // Check if user navigated to /admin or #admin
    const checkAdminRoute = () => {
      const path = window.location.pathname;
      const hash = window.location.hash;
      if (path === '/admin' || hash === '#admin') {
        setIsAdminView(true);
      }
    };
    checkAdminRoute();
    window.addEventListener('hashchange', checkAdminRoute);
    return () => window.removeEventListener('hashchange', checkAdminRoute);
  }, []);

  const loadConfig = async () => {
    try {
      const res = await fetch('/api/admin/config');
      if (res.ok) {
        const data = await res.json();
        setConfig(data);
        setAdminSiteId(data.site_id || '');
        setAdminPod(data.pod || 'p01');
        setAdminUsername(data.username || '');
        setAdminBatchSize(data.batch_size || 50);
        setAdminThrottle(data.throttle_seconds || 2);
        setAdminMaxRows(data.max_row_limit || 5000);
      }
    } catch (err) {
      console.error('Failed to load config:', err);
    }
  };

  const loadCachedForms = async () => {
    try {
      const res = await fetch('/api/forms');
      if (res.ok) {
        const data = await res.json();
        setFormsList(data.forms || []);
        setTotalCachedForms(data.total_cached || 0);
      }
    } catch (err) {
      console.error('Failed to load cached forms:', err);
    }
  };

  // Real Sync Forms call to Eloqua REST API 2.0
  const handleSyncForms = async () => {
    setIsSyncingForms(true);
    setSyncFeedback(null);
    try {
      const res = await fetch('/api/forms/sync', { method: 'POST' });
      const data = await res.json();
      if (res.ok && data.success) {
        setSyncFeedback(`Sync Complete: ${data.synced_count} active forms found.`);
        await loadCachedForms();
        await loadConfig();
      } else {
        setSyncFeedback(`Sync error: ${data.message || 'Failed to sync forms from Eloqua.'}`);
      }
    } catch (err: any) {
      setSyncFeedback(`Network error during sync: ${err.message}`);
    } finally {
      setIsSyncingForms(false);
    }
  };

  // Select form and fetch real field definitions from Eloqua API
  const handleSelectForm = async (form: EloquaForm) => {
    setSelectedForm(form);
    setFormSearch(form.name);
    setFormDropdownOpen(false);
    await fetchFormFields(form.form_id);
  };

  const fetchFormFields = async (formId: number) => {
    setIsRefreshingFields(true);
    try {
      const res = await fetch(`/api/forms/${formId}/fields`);
      const data = await res.json();
      if (res.ok && data.fields) {
        setFormFields(data.fields);
      } else {
        // If live Eloqua has not loaded fields yet, fallback gracefully
        setFormFields([]);
      }
    } catch (err) {
      console.error('Failed to fetch form fields:', err);
      setFormFields([]);
    } finally {
      setIsRefreshingFields(false);
    }
  };

  const handleRefreshFields = () => {
    if (!selectedForm) return;
    fetchFormFields(selectedForm.form_id);
  };

  // Real Test Connection against Eloqua API 2.0
  const handleTestConnection = async () => {
    setAdminTesting(true);
    setAdminTestResult(null);

    try {
      const res = await fetch('/api/admin/test-connection', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          site_id: adminSiteId,
          pod: adminPod,
          username: adminUsername,
          password: adminPassword
        })
      });

      const data = await res.json();
      setAdminTestResult({
        success: data.success,
        message: data.message || (data.success ? 'Connection verified.' : 'Authentication failed.')
      });
    } catch (err: any) {
      setAdminTestResult({
        success: false,
        message: `Connection request failed: ${err.message}`
      });
    } finally {
      setAdminTesting(false);
    }
  };

  // Save Admin Configuration to SQLite
  const handleSaveAdminConfig = async (e: React.FormEvent) => {
    e.preventDefault();
    setAdminSaving(true);

    try {
      const res = await fetch('/api/admin/config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          site_id: adminSiteId,
          pod: adminPod,
          username: adminUsername,
          password: adminPassword,
          batch_size: adminBatchSize,
          throttle_seconds: adminThrottle,
          max_row_limit: adminMaxRows
        })
      });

      const data = await res.json();
      if (res.ok) {
        alert(data.message || 'Configuration saved successfully.');
        setAdminPassword('');
        await loadConfig();
      } else {
        alert(`Failed to save configuration: ${data.message}`);
      }
    } catch (err: any) {
      alert(`Error saving configuration: ${err.message}`);
    } finally {
      setAdminSaving(false);
    }
  };

  // File Upload Handling
  const processFileContent = (file: File) => {
    setUploadError(null);
    setIsUploading(true);

    const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    if (ext !== '.csv' && ext !== '.xlsx') {
      setUploadError('Invalid file format. Please upload a .csv or .xlsx file.');
      setIsUploading(false);
      return;
    }

    const reader = new FileReader();

    if (ext === '.csv') {
      reader.onload = (e) => {
        const text = e.target?.result as string;
        Papa.parse(text, {
          header: true,
          skipEmptyLines: true,
          complete: (results) => {
            const rawHeaders = results.meta.fields || [];
            const rawRows = (results.data as Record<string, any>[]).slice(0, config.max_row_limit);
            const rows = rawRows.map(r => {
              const cleaned: Record<string, any> = {};
              rawHeaders.forEach(h => {
                let v = r[h] !== undefined && r[h] !== null ? r[h] : '';
                if (h.toLowerCase().includes('date') || h.toLowerCase().includes('time') || h.toLowerCase().includes('dob')) {
                  v = normalizeEloquaDate(v);
                }
                cleaned[h] = String(v);
              });
              return cleaned;
            });
            finalizeUpload(file.name, rawHeaders, rows);
          },
          error: (err: Error) => {
            setUploadError(`CSV Parse Error: ${err.message}`);
            setIsUploading(false);
          }
        });
      };
      reader.readAsText(file);
    } else {
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array', cellDates: true });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const jsonRows = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1, raw: false, dateNF: 'mm/dd/yyyy' });

          if (jsonRows.length === 0) {
            setUploadError('The Excel spreadsheet contains no data.');
            setIsUploading(false);
            return;
          }

          const rawHeaders = (jsonRows[0] as any[]).map((h, i) => (h ? String(h).trim() : `Column_${i + 1}`));
          const rows: Record<string, any>[] = [];

          for (let r = 1; r < jsonRows.length; r++) {
            if (rows.length >= config.max_row_limit) break;
            const rowValues = jsonRows[r] as any[];
            if (!rowValues || !rowValues.some(v => v !== undefined && v !== null && String(v).trim() !== '')) {
              continue;
            }
            const rowObj: Record<string, any> = {};
            rawHeaders.forEach((h, hIdx) => {
              let cellVal = rowValues[hIdx] !== undefined && rowValues[hIdx] !== null ? rowValues[hIdx] : '';
              if (cellVal instanceof Date || h.toLowerCase().includes('date') || h.toLowerCase().includes('time') || h.toLowerCase().includes('dob')) {
                cellVal = normalizeEloquaDate(cellVal);
              }
              rowObj[h] = String(cellVal);
            });
            rows.push(rowObj);
          }

          finalizeUpload(file.name, rawHeaders, rows);
        } catch (err: any) {
          setUploadError(`Excel Parsing Error: ${err.message}`);
          setIsUploading(false);
        }
      };
      reader.readAsArrayBuffer(file);
    }
  };

  const finalizeUpload = (filename: string, rawHeaders: string[], rows: Record<string, any>[]) => {
    if (rows.length === 0) {
      setUploadError('Uploaded file contains headers but no valid data rows.');
      setIsUploading(false);
      return;
    }

    const seen = new Set<string>();
    const duplicates: string[] = [];
    rawHeaders.forEach(h => {
      if (seen.has(h)) duplicates.push(h);
      else seen.add(h);
    });

    const newJobId = 'job_' + Math.random().toString(36).substring(2, 10);
    const data: UploadedData = {
      job_id: newJobId,
      filename,
      headers: rawHeaders,
      duplicate_headers: Array.from(new Set(duplicates)),
      total_rows: rows.length,
      sample_row: rows[0] || {},
      rows
    };

    setUploadedData(data);
    setIsUploading(false);
    autoMatchFields(rawHeaders, rows[0] || {});
  };

  const autoMatchFields = (headers: string[], sampleRow: Record<string, any>) => {
    const newMapping: Record<string, EloquaField | "ignore"> = {};
    const usedFieldIds = new Set<number>();

    headers.forEach(header => {
      const norm = header.toLowerCase().trim();
      let matched: EloquaField | null = null;

      // 1. Alias dictionary (avoiding already assigned form fields)
      for (const [alias, targetKey] of Object.entries(ALIAS_MAP)) {
        if (norm === alias || norm.includes(alias)) {
          matched = formFields.find(f => !usedFieldIds.has(f.id) && f.name.toLowerCase().includes(targetKey)) || null;
          if (matched) break;
        }
      }

      // 2. Levenshtein ratio > 0.70 (avoiding already assigned form fields)
      if (!matched) {
        let bestScore = 0;
        for (const f of formFields) {
          if (usedFieldIds.has(f.id)) continue;
          const score = levenshteinSimilarity(header, f.name);
          if (score > bestScore && score >= 0.70) {
            bestScore = score;
            matched = f;
          }
        }
      }

      if (matched) {
        usedFieldIds.add(matched.id);
      }

      newMapping[header] = matched || "ignore";
    });

    setMapping(newMapping);
  };

  const getDuplicateFieldMappings = () => {
    const map = new Map<number, { name: string; cols: string[] }>();
    for (const [col, target] of Object.entries(mapping)) {
      if (target && target !== "ignore") {
        const id = target.id;
        if (!map.has(id)) {
          map.set(id, { name: target.name, cols: [col] });
        } else {
          map.get(id)!.cols.push(col);
        }
      }
    }
    return Array.from(map.values()).filter(item => item.cols.length > 1);
  };

  const isEmailMapped = (): boolean => {
    for (const [col, target] of Object.entries(mapping)) {
      if (target && target !== "ignore") {
        if (target.name.toLowerCase().includes("email") || col.toLowerCase().includes("email")) {
          return true;
        }
      }
    }
    return false;
  };

  // Submit Job to Backend Batch Processor with Circuit Breaker
  const executeBatchProcessing = async () => {
    if (!uploadedData || !selectedForm) return;

    setIsProcessing(true);
    setProcessedRows(0);
    setSuccessfulRows(0);
    setFailedRows(0);
    setCircuitBreakerTriggered(false);
    setCircuitBreakerReason("");

    try {
      const res = await fetch('/api/jobs/submit', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          form_id: selectedForm.form_id,
          form_name: selectedForm.name,
          mapping,
          rows: uploadedData.rows
        })
      });

      const data = await res.json();
      if (!res.ok) {
        alert(data.error || 'Failed to start batch submission.');
        setIsProcessing(false);
        return;
      }

      const jobId = data.job_id;
      setActiveJobId(jobId);

      // Poll job status
      if (pollingRef.current) clearInterval(pollingRef.current);
      pollingRef.current = setInterval(async () => {
        try {
          const statusRes = await fetch(`/api/jobs/${jobId}/status`);
          if (statusRes.ok) {
            const job = await statusRes.json();
            setProcessedRows(job.processed_rows || 0);
            setSuccessfulRows(job.successful_rows || 0);
            setFailedRows(job.failed_rows || 0);

            if (job.status === 'halted_error') {
              if (pollingRef.current) clearInterval(pollingRef.current);
              setIsProcessing(false);
              setCircuitBreakerTriggered(true);
              setCircuitBreakerReason(job.error_summary || '100% failure rate in batch.');
            } else if (job.status === 'completed' || job.status === 'failed') {
              if (pollingRef.current) clearInterval(pollingRef.current);
              setIsProcessing(false);
            }
          }
        } catch (e) {
          console.error('Error polling status:', e);
        }
      }, 1000);
    } catch (err: any) {
      alert(`Submission error: ${err.message}`);
      setIsProcessing(false);
    }
  };

  const base64ToBlob = (base64: string, mimeType: string) => {
    const byteCharacters = atob(base64);
    const byteNumbers = new Array(byteCharacters.length);
    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    const byteArray = new Uint8Array(byteNumbers);
    return new Blob([byteArray], { type: mimeType });
  };

  const triggerBlobDownload = (blob: Blob, filename: string) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  };

  const handleDownloadRetrievedFile = async (format: 'csv' | 'xlsx') => {
    if (!retrieveResult) return;

    if (format === 'csv') {
      const filename = retrieveResult.csv_filename || retrieveResult.filename || 'eloqua_form_data.csv';
      if (retrieveResult.csv_content) {
        const blob = new Blob([retrieveResult.csv_content], { type: 'text/csv;charset=utf-8;' });
        triggerBlobDownload(blob, filename);
      } else {
        try {
          const res = await fetch(retrieveResult.csv_download_url || retrieveResult.download_url);
          if (!res.ok) throw new Error(`HTTP ${res.status}`);
          const blob = await res.blob();
          triggerBlobDownload(blob, filename);
        } catch (e: any) {
          alert('Download error: ' + e.message);
        }
      }
    } else {
      const filename = retrieveResult.xlsx_filename || `${(retrieveResult.filename || 'eloqua_form_data').replace(/\.csv$/, '')}.xlsx`;
      if (retrieveResult.xlsx_base64) {
        const blob = base64ToBlob(retrieveResult.xlsx_base64, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        triggerBlobDownload(blob, filename);
      } else {
        try {
          const res = await fetch(retrieveResult.xlsx_download_url || `${retrieveResult.download_url}&format=xlsx`);
          if (!res.ok) throw new Error(`HTTP ${res.status}`);
          const blob = await res.blob();
          triggerBlobDownload(blob, filename);
        } catch (e: any) {
          alert('Download error: ' + e.message);
        }
      }
    }
  };

  const handleDownloadFailedRecords = async () => {
    if (!activeJobId) return;
    try {
      const res = await fetch(`/api/jobs/${activeJobId}/download-errors`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const blob = await res.blob();
      triggerBlobDownload(blob, `failed_records_${activeJobId}.csv`);
    } catch (err: any) {
      alert(`Failed to download error log: ${err.message}`);
    }
  };

  const handleExecuteRetrieve = async () => {
    if (!retrieveSelectedForm || isRetrieving) return;
    setIsRetrieving(true);
    setRetrieveError(null);
    setRetrieveResult(null);

    try {
      const res = await fetch(`/api/forms/${retrieveSelectedForm.form_id}/export`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          form_id: retrieveSelectedForm.form_id,
          count: retrieveRecordCount
        })
      });

      const contentType = res.headers.get('content-type') || '';
      if (!res.ok || !contentType.includes('application/json')) {
        let errText = await res.text();
        try {
          const errJson = JSON.parse(errText);
          errText = errJson.message || errText;
        } catch {}
        throw new Error(errText || `Server returned HTTP ${res.status}`);
      }

      const data = await res.json();
      if (!data.success) {
        throw new Error(data.message || 'Failed to retrieve form data.');
      }

      setRetrieveResult(data);
    } catch (err: any) {
      console.error('Retrieve error:', err);
      setRetrieveError(err.message || 'An error occurred while retrieving data from Eloqua.');
    } finally {
      setIsRetrieving(false);
    }
  };

  const resetAll = () => {
    setCurrentStep(1);
    setUploadedData(null);
    setMapping({});
    setCircuitBreakerTriggered(false);
    setCircuitBreakerReason("");
    setActiveJobId(null);
    setProcessedRows(0);
    setSuccessfulRows(0);
    setFailedRows(0);
    setRetrieveSelectedForm(null);
    setRetrieveFormSearch('');
    setRetrieveResult(null);
    setRetrieveError(null);
    setAppMode('select');
  };

  const activeMappedCount = Object.values(mapping).filter(v => v !== "ignore").length;
  const ignoredCount = Object.values(mapping).filter(v => v === "ignore").length;

  return (
    <div className="min-h-screen flex flex-col bg-[#F3F5F9] text-slate-800 font-sans selection:bg-blue-100">

      {/* Top Navigation Bar: Strictly matches spec with Logo and Start Over button */}
      <header className="bg-white border-b border-gray-200 sticky top-0 z-40 shadow-xs">
        <div className="max-w-6xl mx-auto px-6 h-16 flex items-center justify-between">
          <div className="flex items-center space-x-3 cursor-pointer" onClick={resetAll}>
            <div className="w-9 h-9 rounded-xl bg-blue-600 flex items-center justify-center text-white shadow-sm shadow-blue-500/25">
              <FileSpreadsheet className="w-5 h-5" />
            </div>
            <h1 className="text-base font-bold text-slate-900 tracking-tight">Eloqua FormOps</h1>
          </div>

          <div className="flex items-center space-x-3">
            {/* Start Over Button */}
            <button
              onClick={resetAll}
              className="inline-flex items-center space-x-1.5 px-3 py-1.5 text-xs font-semibold text-slate-600 hover:text-slate-900 bg-gray-100 hover:bg-gray-200 rounded-lg transition-colors cursor-pointer"
              title="Reset and return to operation selection"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              <span>Start Over</span>
            </button>
          </div>
        </div>
      </header>

      {/* Admin Panel (Hidden route accessible strictly via /admin or #admin) */}
      {isAdminView ? (
        <main className="flex-1 max-w-2xl w-full mx-auto p-6 mb-20 animate-in fade-in duration-200">
          <div className="bg-white rounded-2xl border border-gray-200 shadow-sm p-8">
            <div className="flex items-center justify-between pb-6 border-b border-gray-100">
              <div>
                <h2 className="text-xl font-bold text-slate-900 flex items-center space-x-2">
                  <Sliders className="w-5 h-5 text-blue-600" />
                  <span>System Configuration</span>
                </h2>
                <p className="text-xs text-slate-500 mt-1">
                  Oracle Eloqua REST API 2.0 Credentials, Batch Size & Throttle Limits
                </p>
              </div>
              <span className="px-2.5 py-1 bg-amber-50 text-amber-800 border border-amber-200 rounded-full text-xs font-semibold">
                Hidden Route (/admin)
              </span>
            </div>

            <form onSubmit={handleSaveAdminConfig} className="space-y-5 mt-6">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-bold uppercase text-slate-600 mb-1">Eloqua Pod</label>
                  <input
                    type="text"
                    value={adminPod}
                    onChange={e => setAdminPod(e.target.value)}
                    required
                    className="w-full px-3.5 py-2 text-sm bg-slate-50 border border-gray-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-blue-500 outline-none"
                    placeholder="p01"
                  />
                  <p className="text-[11px] text-slate-400 mt-1">e.g. p01, p02</p>
                </div>
                <div>
                  <label className="block text-xs font-bold uppercase text-slate-600 mb-1">Eloqua Company Name</label>
                  <input
                    type="text"
                    value={adminSiteId}
                    onChange={e => setAdminSiteId(e.target.value)}
                    required
                    className="w-full px-3.5 py-2 text-sm bg-slate-50 border border-gray-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-blue-500 outline-none"
                    placeholder="e.g. company"
                  />
                  <p className="text-[11px] text-slate-400 mt-1">Company name for authentication</p>
                </div>
              </div>

              <div>
                <label className="block text-xs font-bold uppercase text-slate-600 mb-1">Eloqua Username</label>
                <input
                  type="text"
                  value={adminUsername}
                  onChange={e => setAdminUsername(e.target.value)}
                  className="w-full px-3.5 py-2 text-sm bg-slate-50 border border-gray-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-blue-500 outline-none"
                  placeholder="e.g. first.last"
                />
              </div>

              <div>
                <label className="block text-xs font-bold uppercase text-slate-600 mb-1">Password</label>
                <input
                  type="password"
                  value={adminPassword}
                  onChange={e => setAdminPassword(e.target.value)}
                  placeholder={config.has_password ? "Password saved in DB. Leave blank to preserve." : "Enter Eloqua API password"}
                  className="w-full px-3.5 py-2 text-sm bg-slate-50 border border-gray-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-blue-500 outline-none"
                />
                <p className="text-[11px] text-slate-400 mt-1 flex items-center space-x-1">
                  <ShieldCheck className="w-3.5 h-3.5 text-emerald-600 inline" />
                  <span>Zero Plaintext: Stored in SQLite storage.db and never displayed or returned in the UI.</span>
                </p>
              </div>

              <div className="grid grid-cols-3 gap-4 pt-2">
                <div>
                  <label className="block text-xs font-bold uppercase text-slate-600 mb-1">Batch Size</label>
                  <input
                    type="number"
                    value={adminBatchSize}
                    onChange={e => setAdminBatchSize(Number(e.target.value))}
                    min={1}
                    max={500}
                    required
                    className="w-full px-3 py-2 text-sm bg-slate-50 border border-gray-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-blue-500 outline-none"
                  />
                  <p className="text-[11px] text-slate-400 mt-1">Default: 50 records</p>
                </div>
                <div>
                  <label className="block text-xs font-bold uppercase text-slate-600 mb-1">Throttle (Sec)</label>
                  <input
                    type="number"
                    step="0.5"
                    value={adminThrottle}
                    onChange={e => setAdminThrottle(Number(e.target.value))}
                    min={0}
                    max={30}
                    required
                    className="w-full px-3 py-2 text-sm bg-slate-50 border border-gray-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-blue-500 outline-none"
                  />
                  <p className="text-[11px] text-slate-400 mt-1">Default: 2.0s</p>
                </div>
                <div>
                  <label className="block text-xs font-bold uppercase text-slate-600 mb-1">Max Rows</label>
                  <input
                    type="number"
                    value={adminMaxRows}
                    onChange={e => setAdminMaxRows(Number(e.target.value))}
                    min={10}
                    max={20000}
                    required
                    className="w-full px-3 py-2 text-sm bg-slate-50 border border-gray-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-blue-500 outline-none"
                  />
                  <p className="text-[11px] text-slate-400 mt-1">Default: 5,000</p>
                </div>
              </div>

              {adminTestResult && (
                <div
                  className={`p-3.5 rounded-xl text-xs font-medium border ${
                    adminTestResult.success
                      ? 'bg-emerald-50 text-emerald-800 border-emerald-200'
                      : 'bg-rose-50 text-rose-800 border-rose-200'
                  }`}
                >
                  {adminTestResult.success ? '✓ ' : '✕ '}
                  {adminTestResult.message}
                </div>
              )}

              <div className="pt-6 border-t border-gray-100 flex items-center justify-between">
                <button
                  type="button"
                  onClick={handleTestConnection}
                  disabled={adminTesting}
                  className="px-4 py-2 text-xs font-bold text-slate-700 bg-gray-100 hover:bg-gray-200 rounded-xl transition-colors disabled:opacity-50"
                >
                  {adminTesting ? 'Connecting to Eloqua...' : 'Test Connection'}
                </button>
                <div className="flex space-x-2">
                  <button
                    type="button"
                    onClick={() => {
                      setIsAdminView(false);
                      window.location.hash = '';
                    }}
                    className="px-4 py-2 text-xs font-semibold text-slate-600 hover:text-slate-800 rounded-xl"
                  >
                    Close
                  </button>
                  <button
                    type="submit"
                    disabled={adminSaving}
                    className="px-5 py-2 text-xs font-bold text-white bg-blue-600 hover:bg-blue-700 rounded-xl shadow-sm transition-colors"
                  >
                    {adminSaving ? 'Saving...' : 'Save Configuration'}
                  </button>
                </div>
              </div>
            </form>
          </div>
        </main>
      ) : appMode === 'select' ? (
        <main className="flex-1 max-w-4xl w-full mx-auto p-6 my-auto flex flex-col justify-center animate-in fade-in">
          <div className="bg-white rounded-3xl border border-gray-200 p-8 sm:p-14 shadow-sm text-center">
            <div className="max-w-2xl mx-auto mb-10">
              <h2 className="text-2xl sm:text-3xl font-extrabold text-slate-900 tracking-tight">Choose Operation</h2>
              <p className="text-sm text-slate-500 mt-2">Submit spreadsheet records in batches to an Eloqua form, or extract recent form submissions directly to a file.</p>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 max-w-3xl mx-auto text-left">
              <button
                onClick={() => { setAppMode('submit'); setCurrentStep(1); }}
                type="button"
                className="group relative p-8 bg-white hover:bg-blue-50/50 rounded-2xl border-2 border-blue-600 hover:border-blue-700 shadow-sm hover:shadow-md transition-all text-left flex flex-col justify-between cursor-pointer focus:outline-none focus:ring-4 focus:ring-blue-100 min-h-[220px]"
              >
                <div>
                  <div className="w-12 h-12 rounded-xl bg-blue-600 text-white flex items-center justify-center shadow-md shadow-blue-500/20 mb-6 group-hover:scale-105 transition-transform">
                    <UploadCloud className="w-6 h-6" />
                  </div>
                  <h3 className="text-lg font-bold text-slate-900 group-hover:text-blue-600 transition-colors">Submit form data to Eloqua</h3>
                  <p className="text-xs text-slate-500 mt-2 leading-relaxed">
                    Upload Excel or CSV spreadsheets, intelligently map columns to form fields with fuzzy matching, and submit records in controlled batches with circuit breaker protection.
                  </p>
                </div>
                <div className="mt-6 flex items-center text-xs font-bold text-blue-600 group-hover:translate-x-1 transition-transform">
                  <span>Start Batch Ingestion</span>
                  <ChevronRight className="w-4 h-4 ml-1.5" />
                </div>
              </button>

              <button
                onClick={() => setAppMode('retrieve')}
                type="button"
                className="group relative p-8 bg-white hover:bg-blue-50/50 rounded-2xl border-2 border-blue-600 hover:border-blue-700 shadow-sm hover:shadow-md transition-all text-left flex flex-col justify-between cursor-pointer focus:outline-none focus:ring-4 focus:ring-blue-100 min-h-[220px]"
              >
                <div>
                  <div className="w-12 h-12 rounded-xl bg-blue-600 text-white flex items-center justify-center shadow-md shadow-blue-500/20 mb-6 group-hover:scale-105 transition-transform">
                    <Download className="w-6 h-6" />
                  </div>
                  <h3 className="text-lg font-bold text-slate-900 group-hover:text-blue-600 transition-colors">Retrieve form data from Eloqua</h3>
                  <p className="text-xs text-slate-500 mt-2 leading-relaxed">
                    Select any synced Eloqua form, choose how many recent submissions to pull (10, 100, 500, or 1000), and download the assembled JSON data as a CSV or Excel file.
                  </p>
                </div>
                <div className="mt-6 flex items-center text-xs font-bold text-blue-600 group-hover:translate-x-1 transition-transform">
                  <span>Start Data Retrieval</span>
                  <ChevronRight className="w-4 h-4 ml-1.5" />
                </div>
              </button>
            </div>
          </div>
        </main>
      ) : appMode === 'retrieve' ? (
        <main className="flex-1 max-w-4xl w-full mx-auto p-6 mb-20 animate-in fade-in">
          <div className="bg-white rounded-2xl border border-gray-200 p-8 shadow-xs space-y-8">
            <div className="flex flex-col sm:flex-row sm:items-start justify-between gap-4 pb-6 border-b border-gray-100">
              <div>
                <h2 className="text-xl font-bold text-slate-900">Retrieve Form Data from Eloqua</h2>
                <p className="text-sm text-slate-500 mt-1">Select an Eloqua form and choose how many recent submission records to export into a file.</p>
              </div>
              <div className="flex space-x-2">
                <button
                  onClick={handleSyncForms}
                  disabled={isSyncingForms}
                  className="inline-flex items-center space-x-1.5 px-3 py-1.5 text-xs font-semibold text-blue-700 bg-blue-50 hover:bg-blue-100 rounded-lg border border-blue-200 transition-colors cursor-pointer"
                >
                  <RefreshCw className={`w-3.5 h-3.5 ${isSyncingForms ? 'animate-spin' : ''}`} />
                  <span>{isSyncingForms ? 'Syncing...' : 'Sync Forms'}</span>
                </button>
              </div>
            </div>

            {syncFeedback && (
              <div className="p-3 bg-blue-50 border border-blue-200 rounded-xl text-xs text-blue-900">
                {syncFeedback}
              </div>
            )}

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="relative">
                <label className="block text-xs font-bold uppercase tracking-wider text-slate-700 mb-2">
                  1. Target Eloqua Form
                </label>
                <div className="relative">
                  <input
                    type="text"
                    value={retrieveFormSearch}
                    placeholder="Type form name or ID to filter..."
                    onChange={e => {
                      setRetrieveFormSearch(e.target.value);
                      setRetrieveDropdownOpen(true);
                    }}
                    onFocus={() => setRetrieveDropdownOpen(true)}
                    className="w-full px-4 py-3 pl-10 text-sm bg-slate-50 border border-gray-300 rounded-xl focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 transition-all"
                  />
                  <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-3.5" />
                </div>

                {retrieveDropdownOpen && (
                  <div className="absolute top-full left-0 right-0 mt-1 max-h-72 overflow-y-auto bg-white border border-gray-200 rounded-xl shadow-xl z-50 divide-y divide-gray-100">
                    {formsList
                      .filter(f => !retrieveFormSearch.trim() || f.name.toLowerCase().includes(retrieveFormSearch.toLowerCase()) || String(f.form_id).includes(retrieveFormSearch))
                      .slice(0, 50)
                      .map(form => (
                        <div
                          key={form.form_id}
                          onClick={() => {
                            setRetrieveSelectedForm(form);
                            setRetrieveFormSearch(form.name);
                            setRetrieveDropdownOpen(false);
                            setRetrieveResult(null);
                            setRetrieveError(null);
                          }}
                          className="p-3 hover:bg-blue-50 cursor-pointer flex justify-between items-center text-xs transition-colors"
                        >
                          <span className="font-semibold text-slate-800 truncate pr-3">{form.name}</span>
                          <span className="text-slate-400 font-mono text-[11px] bg-slate-100 px-2 py-0.5 rounded">ID: {form.form_id}</span>
                        </div>
                      ))}
                  </div>
                )}

                {retrieveSelectedForm && (
                  <div className="mt-3 p-3.5 bg-blue-50 border border-blue-200 rounded-xl flex items-center justify-between">
                    <div className="flex items-center space-x-3 truncate">
                      <div className="w-8 h-8 rounded-lg bg-blue-600 text-white flex items-center justify-center font-bold text-[11px] shrink-0">ID</div>
                      <div className="truncate">
                        <p className="text-xs font-bold text-blue-950 truncate">{retrieveSelectedForm.name}</p>
                        <p className="text-[11px] text-blue-700 truncate">Eloqua Numeric ID: {retrieveSelectedForm.form_id}</p>
                      </div>
                    </div>
                    <button
                      onClick={() => {
                        setRetrieveSelectedForm(null);
                        setRetrieveFormSearch('');
                      }}
                      className="text-xs text-blue-700 hover:text-blue-900 font-semibold ml-2 cursor-pointer"
                    >
                      Change
                    </button>
                  </div>
                )}
                <p className="text-[11px] text-slate-400 mt-2">{totalCachedForms} forms in SQLite storage.db</p>
              </div>

              <div>
                <label className="block text-xs font-bold uppercase tracking-wider text-slate-700 mb-2">
                  2. Number of Recent Records to Pull
                </label>
                <div className="relative">
                  <select
                    value={retrieveRecordCount}
                    onChange={e => setRetrieveRecordCount(Number(e.target.value))}
                    className="w-full px-4 py-3 text-sm bg-slate-50 border border-gray-300 rounded-xl focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 transition-all font-medium text-slate-800"
                  >
                    <option value={10}>10 records (Quick sample)</option>
                    <option value={100}>100 records (Recommended)</option>
                    <option value={500}>500 records</option>
                    <option value={1000}>1000 records (Maximum)</option>
                  </select>
                </div>
                <p className="text-[11px] text-slate-400 mt-2">
                  Pulls the most recent submissions based on submission date from Oracle Eloqua.
                </p>
              </div>
            </div>

            <div className="pt-4 border-t border-gray-100 flex items-center justify-end">
              <button
                onClick={handleExecuteRetrieve}
                disabled={!retrieveSelectedForm || isRetrieving}
                className="inline-flex items-center space-x-2 px-6 py-3 bg-blue-600 hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed text-white text-xs font-bold rounded-xl shadow-md transition-all cursor-pointer"
              >
                <Download className="w-4 h-4" />
                <span>{isRetrieving ? 'Retrieving Data...' : 'Retrieve Form Data'}</span>
              </button>
            </div>

            {isRetrieving && (
              <div className="p-6 bg-slate-50 border border-slate-200 rounded-xl text-center space-y-3">
                <div className="inline-block animate-spin w-8 h-8 border-3 border-blue-600 border-t-transparent rounded-full" />
                <h4 className="text-sm font-bold text-slate-800">Retrieving form data from Oracle Eloqua...</h4>
                <p className="text-xs text-slate-500">Fetching records via Eloqua REST API 2.0 and mapping field identifiers...</p>
              </div>
            )}

            {retrieveError && (
              <div className="p-4 bg-rose-50 border border-rose-200 rounded-xl text-xs text-rose-900 space-y-1">
                <div className="font-bold text-sm text-rose-950 flex items-center space-x-2">
                  <AlertCircle className="w-4 h-4 text-rose-600" />
                  <span>Retrieval Failed</span>
                </div>
                <p>{retrieveError}</p>
              </div>
            )}

            {retrieveResult && (
              <div className="space-y-6 pt-2">
                <div className="p-6 bg-emerald-50 border border-emerald-200 rounded-2xl flex flex-col sm:flex-row items-center justify-between gap-4">
                  <div className="flex items-center space-x-3.5">
                    <div className="w-12 h-12 bg-emerald-600 text-white rounded-xl flex items-center justify-center shrink-0 shadow-sm">
                      <CheckCircle2 className="w-6 h-6" />
                    </div>
                    <div>
                      <h3 className="text-base font-bold text-emerald-950">
                        {retrieveResult.total_records === 0 ? '0 Submissions Found' : 'Data Retrieved Successfully!'}
                      </h3>
                      <p className="text-xs text-emerald-700 mt-0.5">
                        {retrieveResult.total_records === 0
                          ? `Form: ${retrieveResult.form_name} (ID: ${retrieveResult.form_id}) has 0 recorded submissions in Oracle Eloqua.`
                          : `Form: ${retrieveResult.form_name} (ID: ${retrieveResult.form_id}) • ${retrieveResult.total_records} records retrieved`}
                      </p>
                    </div>
                  </div>
                  <div className="flex flex-wrap items-center gap-2.5">
                    <button
                      type="button"
                      onClick={() => handleDownloadRetrievedFile('csv')}
                      className="inline-flex items-center space-x-2 px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-xl shadow-md transition-colors cursor-pointer"
                    >
                      <Download className="w-4 h-4" />
                      <span>Download CSV File (.csv)</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => handleDownloadRetrievedFile('xlsx')}
                      className="inline-flex items-center space-x-2 px-4 py-2.5 bg-white hover:bg-emerald-50 text-emerald-800 border border-emerald-300 text-xs font-bold rounded-xl shadow-xs transition-colors cursor-pointer"
                    >
                      <Download className="w-4 h-4 text-emerald-600" />
                      <span>Download Excel (.xlsx)</span>
                    </button>
                  </div>
                </div>

                <div className="flex flex-col sm:flex-row sm:items-center justify-between text-[11px] text-slate-500 gap-1 px-1">
                  <span className="font-medium">
                    {retrieveResult.csv_filename ? `CSV: ${retrieveResult.csv_filename} | Excel: ${retrieveResult.xlsx_filename}` : `File: ${retrieveResult.filename}`} • {retrieveResult.columns.length} columns mapped
                  </span>
                  <span className="text-slate-400">Temporary file is automatically purged upon download or new export request</span>
                </div>

                <div className="space-y-2">
                  <div className="flex items-center justify-between">
                    <h4 className="text-xs font-bold uppercase tracking-wider text-slate-700">Preview Top Records</h4>
                    <span className="text-xs text-slate-400">Showing first {retrieveResult.preview_rows.length} of {retrieveResult.total_records} records</span>
                  </div>
                  <div className="border border-gray-200 rounded-xl overflow-hidden shadow-xs">
                    <div className="overflow-x-auto max-h-96">
                      <table className="w-full text-left text-xs border-collapse">
                        <thead className="bg-slate-100 text-slate-700 uppercase font-bold sticky top-0 z-10 border-b border-gray-200">
                          <tr>
                            {retrieveResult.columns.map(col => (
                              <th key={col} className="px-3.5 py-2.5 font-bold whitespace-nowrap text-slate-700 bg-slate-100 border-b border-gray-200">
                                {col}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-100 bg-white">
                          {retrieveResult.preview_rows.length === 0 ? (
                            <tr>
                              <td colSpan={retrieveResult.columns.length} className="px-4 py-8 text-center text-slate-400">
                                No submissions found for this form.
                              </td>
                            </tr>
                          ) : (
                            retrieveResult.preview_rows.map((row, rIdx) => (
                              <tr key={rIdx} className="hover:bg-slate-50 transition-colors">
                                {retrieveResult.columns.map(col => (
                                  <td key={col} className="px-3.5 py-2 font-normal text-slate-700 whitespace-nowrap max-w-xs truncate" title={String(row[col] ?? '')}>
                                    {String(row[col] ?? '') || <span className="text-slate-300 italic">empty</span>}
                                  </td>
                                ))}
                              </tr>
                            ))
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>

                <div className="pt-4 flex justify-between items-center border-t border-gray-100">
                  <button
                    onClick={() => {
                      setRetrieveSelectedForm(null);
                      setRetrieveFormSearch('');
                      setRetrieveResult(null);
                      setRetrieveError(null);
                    }}
                    className="px-4 py-2 text-xs font-semibold text-slate-600 hover:text-slate-800 bg-gray-100 hover:bg-gray-200 rounded-xl transition-colors cursor-pointer"
                  >
                    Retrieve Another Form
                  </button>
                  <button
                    onClick={resetAll}
                    className="px-4 py-2 text-xs font-semibold text-blue-600 hover:text-blue-800 rounded-xl cursor-pointer"
                  >
                    ← Return to Operation Selection
                  </button>
                </div>
              </div>
            )}
          </div>
        </main>
      ) : (
        <>
          {/* Step Tracker Progress Bar */}
          <div className="bg-white border-b border-gray-200 py-3 shadow-xs">
            <div className="max-w-4xl mx-auto px-6">
              <div className="flex items-center justify-between text-xs font-medium">
                {[
                  { step: 1, label: "Select Form" },
                  { step: 2, label: "Upload File" },
                  { step: 3, label: "Map Fields" },
                  { step: 4, label: "Review Mapping" },
                  { step: 5, label: "Process & Results" }
                ].map((item, idx, arr) => (
                  <React.Fragment key={item.step}>
                    <div
                      className={`flex items-center space-x-2 ${
                        currentStep === item.step
                          ? 'text-blue-600 font-bold'
                          : currentStep > item.step
                          ? 'text-emerald-600'
                          : 'text-slate-400'
                      }`}
                    >
                      <span
                        className={`w-6 h-6 rounded-full border-2 flex items-center justify-center font-bold text-xs ${
                          currentStep === item.step
                            ? 'border-blue-600 bg-blue-50 text-blue-600'
                            : currentStep > item.step
                            ? 'border-emerald-600 bg-emerald-50 text-emerald-600'
                            : 'border-gray-200 text-gray-400'
                        }`}
                      >
                        {currentStep > item.step ? '✓' : item.step}
                      </span>
                      <span className="hidden sm:inline">{item.label}</span>
                    </div>
                    {idx < arr.length - 1 && (
                      <div className="flex-1 h-0.5 bg-gray-200 mx-3"></div>
                    )}
                  </React.Fragment>
                ))}
              </div>
            </div>
          </div>

          {/* Main Wizard Container */}
          <main className="flex-1 max-w-4xl w-full mx-auto p-6 mb-20">

            {/* STEP 1: Select Form */}
            {currentStep === 1 && (
              <section className="space-y-6 animate-in fade-in duration-150">
                <div className="bg-white rounded-2xl border border-gray-200 p-8 shadow-xs">
                  <div className="flex flex-col sm:flex-row sm:items-start justify-between gap-4">
                    <div>
                      <h2 className="text-xl font-bold text-slate-900">Step 1: Choose Target Eloqua Form</h2>
                      <p className="text-sm text-slate-500 mt-1">
                        Select a form to retrieve its schema and numeric field identifiers from Oracle Eloqua.
                      </p>
                    </div>
                    <div className="flex space-x-2">
                      <button
                        onClick={handleSyncForms}
                        disabled={isSyncingForms}
                        className="inline-flex items-center space-x-1.5 px-3 py-1.5 text-xs font-semibold text-blue-700 bg-blue-50 hover:bg-blue-100 rounded-lg border border-blue-200 transition-colors"
                      >
                        <RefreshCw className={`w-3.5 h-3.5 ${isSyncingForms ? 'animate-spin' : ''}`} />
                        <span>{isSyncingForms ? 'Syncing Eloqua...' : 'Sync Forms'}</span>
                      </button>
                      <button
                        onClick={handleRefreshFields}
                        disabled={!selectedForm || isRefreshingFields}
                        className="inline-flex items-center space-x-1.5 px-3 py-1.5 text-xs font-semibold text-slate-700 bg-gray-50 hover:bg-gray-100 rounded-lg border border-gray-200 transition-colors disabled:opacity-50"
                      >
                        <RefreshCw className={`w-3.5 h-3.5 ${isRefreshingFields ? 'animate-spin' : ''}`} />
                        <span>{isRefreshingFields ? 'Refreshing...' : 'Refresh Fields'}</span>
                      </button>
                    </div>
                  </div>

                  {syncFeedback && (
                    <div className="mt-4 p-3 bg-blue-50 border border-blue-200 rounded-xl text-xs text-blue-900">
                      {syncFeedback}
                    </div>
                  )}

                  {/* Form Typeahead Search */}
                  <div className="mt-6 relative">
                    <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-2">
                      Search Form by Name or Numeric ID
                    </label>
                    <div className="relative">
                      <input
                        type="text"
                        value={formSearch}
                        onChange={e => {
                          setFormSearch(e.target.value);
                          setFormDropdownOpen(true);
                        }}
                        onFocus={() => setFormDropdownOpen(true)}
                        placeholder={totalCachedForms > 0 ? "Type to search cached Eloqua forms..." : "No forms in cache. Click 'Sync Forms' or configure /admin..."}
                        className="w-full px-4 py-3 pl-10 text-sm bg-slate-50 border border-gray-300 rounded-xl focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition-all"
                      />
                      <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-3.5" />
                    </div>

                    {formDropdownOpen && (
                      <div className="absolute top-full left-0 right-0 mt-1 max-h-60 overflow-y-auto bg-white border border-gray-200 rounded-xl shadow-lg z-50 divide-y divide-gray-100">
                        {formsList
                          .filter(f =>
                            f.name.toLowerCase().includes(formSearch.toLowerCase()) ||
                            String(f.form_id).includes(formSearch)
                          )
                          .map(f => (
                            <div
                              key={f.form_id}
                              onClick={() => handleSelectForm(f)}
                              className="p-3 hover:bg-blue-50 cursor-pointer flex justify-between items-center text-xs"
                            >
                              <span className="font-semibold text-slate-800">{f.name}</span>
                              <span className="text-slate-400 font-mono">ID: {f.form_id}</span>
                            </div>
                          ))}
                        {formsList.length === 0 && (
                          <div className="p-4 text-xs text-slate-500 text-center">
                            No forms found. Click <strong>Sync Forms</strong> above to fetch your active forms from Oracle Eloqua.
                          </div>
                        )}
                      </div>
                    )}
                  </div>

                  {/* Selected Form Banner */}
                  {selectedForm && (
                    <div className="mt-6 p-4 bg-blue-50 border border-blue-200 rounded-xl flex items-center justify-between">
                      <div className="flex items-center space-x-3">
                        <div className="w-10 h-10 rounded-lg bg-blue-600 text-white flex items-center justify-center font-bold text-xs shadow-xs">
                          ID
                        </div>
                        <div>
                          <p className="text-sm font-bold text-blue-950">{selectedForm.name}</p>
                          <p className="text-xs text-blue-700">
                            Eloqua Form ID: <span className="font-mono font-semibold">{selectedForm.form_id}</span> • {formFields.length} actionable input fields loaded
                          </p>
                        </div>
                      </div>
                      <span className="text-xs font-semibold px-2.5 py-1 bg-blue-200/60 text-blue-900 rounded-md">
                        Selected
                      </span>
                    </div>
                  )}

                  <div className="mt-6 pt-4 border-t border-gray-100 flex items-center justify-between text-xs text-slate-400">
                    <span>{totalCachedForms} forms in SQLite storage.db</span>
                    <span>Last synced: {config.forms_last_synced || 'Never'}</span>
                  </div>
                </div>
              </section>
            )}

            {/* STEP 2: Upload File */}
            {currentStep === 2 && (
              <section className="space-y-6 animate-in fade-in duration-150">
                <div className="bg-white rounded-2xl border border-gray-200 p-8 shadow-xs">
                  <div>
                    <h2 className="text-xl font-bold text-slate-900">Step 2: Upload Spreadsheet</h2>
                    <p className="text-sm text-slate-500 mt-1">
                      Upload your dataset in Excel (.xlsx) or CSV format. Reads the first sheet only.
                    </p>
                  </div>

                  {/* Drag and Drop Zone */}
                  <label
                    onDragOver={e => e.preventDefault()}
                    onDrop={e => {
                      e.preventDefault();
                      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
                        processFileContent(e.dataTransfer.files[0]);
                      }
                    }}
                    className="mt-6 border-2 border-dashed border-gray-300 hover:border-blue-500 bg-slate-50 hover:bg-blue-50/40 rounded-2xl p-12 text-center cursor-pointer transition-all block"
                  >
                    <input
                      type="file"
                      className="hidden"
                      accept=".csv, .xlsx"
                      onChange={e => {
                        if (e.target.files && e.target.files[0]) {
                          processFileContent(e.target.files[0]);
                        }
                      }}
                    />
                    <div className="w-16 h-16 mx-auto mb-4 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center">
                      <UploadCloud className="w-8 h-8" />
                    </div>
                    <h3 className="text-base font-bold text-slate-800">Drag & drop your file here</h3>
                    <p className="text-xs text-slate-500 mt-1">or click anywhere to browse for a .xlsx or .csv file</p>
                    <div className="mt-4 inline-flex items-center space-x-2 px-3 py-1 bg-white border border-gray-200 rounded-full text-xs font-medium text-slate-600">
                      <span>Maximum {config.max_row_limit.toLocaleString()} rows ceiling</span>
                    </div>
                  </label>

                  {uploadError && (
                    <div className="mt-4 p-4 bg-rose-50 border border-rose-200 rounded-xl text-xs text-rose-800 flex items-center space-x-2">
                      <AlertCircle className="w-4 h-4 text-rose-600 flex-shrink-0" />
                      <span>{uploadError}</span>
                    </div>
                  )}

                  {/* File Upload Preview */}
                  {uploadedData && (
                    <div className="mt-6 p-4 bg-emerald-50 border border-emerald-200 rounded-xl flex items-center justify-between">
                      <div className="flex items-center space-x-3">
                        <div className="w-10 h-10 rounded-lg bg-emerald-600 text-white flex items-center justify-center">
                          <CheckCircle2 className="w-5 h-5" />
                        </div>
                        <div>
                          <p className="text-sm font-bold text-emerald-950">{uploadedData.filename}</p>
                          <p className="text-xs text-emerald-700">
                            {uploadedData.total_rows} rows parsed • {uploadedData.headers.length} columns detected (First tab)
                          </p>
                        </div>
                      </div>
                      <button
                        onClick={() => setUploadedData(null)}
                        className="text-xs text-red-600 hover:text-red-800 font-semibold"
                      >
                        Change File
                      </button>
                    </div>
                  )}

                  {/* Duplicate Headers Warning Banner */}
                  {uploadedData && uploadedData.duplicate_headers.length > 0 && (
                    <div className="mt-4 p-4 bg-amber-50 border border-amber-300 rounded-xl text-amber-900 text-xs flex items-start space-x-2">
                      <AlertTriangle className="w-4 h-4 text-amber-600 flex-shrink-0 mt-0.5" />
                      <div>
                        <span className="font-bold">Duplicate Column Headers Detected: </span>
                        <span>{uploadedData.duplicate_headers.join(', ')}. Please verify each mapping carefully.</span>
                      </div>
                    </div>
                  )}
                </div>
              </section>
            )}

            {/* STEP 3: Map Fields */}
            {currentStep === 3 && (
              <section className="space-y-6 animate-in fade-in duration-150">
                <div className="bg-white rounded-2xl border border-gray-200 p-8 shadow-xs">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                    <div>
                      <h2 className="text-xl font-bold text-slate-900">Step 3: Map Columns to Eloqua Fields</h2>
                      <p className="text-sm text-slate-500 mt-1">
                        Headers auto-matched via Levenshtein fuzzy distance & aliases. Sample values shown from Row 1.
                      </p>
                    </div>

                    {/* Email validation badge */}
                    <div className="flex items-center">
                      {isEmailMapped() ? (
                        <span className="inline-flex items-center space-x-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-emerald-100 text-emerald-800 border border-emerald-200">
                          <Check className="w-3.5 h-3.5" />
                          <span>Valid Email Field Mapped</span>
                        </span>
                      ) : (
                        <span className="inline-flex items-center space-x-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-amber-100 text-amber-800 border border-amber-300">
                          <AlertTriangle className="w-3.5 h-3.5" />
                          <span>Email mapping required</span>
                        </span>
                      )}
                    </div>
                  </div>

                  {/* Duplicate Mapping Notice to guarantee compliance with Eloqua NoDuplicatesRequirement */}
                  {getDuplicateFieldMappings().length > 0 && (
                    <div className="mt-4 p-3.5 bg-amber-50 border border-amber-200 rounded-xl text-xs text-amber-900 flex items-start space-x-2">
                      <AlertTriangle className="w-4 h-4 text-amber-600 flex-shrink-0 mt-0.5" />
                      <div>
                        <span className="font-bold">Duplicate Target Field Notice: </span>
                        <span>
                          Multiple columns are mapped to the same Eloqua field ({getDuplicateFieldMappings().map(d => `"${d.name}" [${d.cols.join(', ')}]`).join('; ')}).
                          To satisfy Oracle Eloqua's <code className="font-mono bg-amber-100 px-1 py-0.5 rounded">NoDuplicatesRequirement</code>, the submission engine automatically consolidates non-empty values into a single unique field value.
                        </span>
                      </div>
                    </div>
                  )}

                  {/* Mapping Table */}
                  <div className="mt-6 border border-gray-200 rounded-xl overflow-hidden">
                    <table className="w-full text-left text-xs">
                      <thead className="bg-slate-50 text-slate-500 uppercase tracking-wider font-semibold border-b border-gray-200">
                        <tr>
                          <th className="py-3 px-4">Uploaded Column Header</th>
                          <th className="py-3 px-4">Row 1 Sample Preview</th>
                          <th className="py-3 px-4">Target Eloqua Form Field</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-100 text-slate-700 font-medium">
                        {uploadedData?.headers.map((col, idx) => {
                          const sampleVal = uploadedData.sample_row[col] !== undefined ? String(uploadedData.sample_row[col]) : '';
                          const isDuplicate = uploadedData.duplicate_headers.includes(col);
                          const currentTarget = mapping[col];
                          const selectedValue = currentTarget === "ignore" ? "ignore" : currentTarget?.id ?? "ignore";

                          return (
                            <tr key={idx} className={isDuplicate ? 'bg-amber-50/70' : ''}>
                              <td className="py-3 px-4 font-semibold">
                                <div className="flex items-center space-x-2">
                                  <span>{col}</span>
                                  {isDuplicate && (
                                    <span className="px-1.5 py-0.5 text-[10px] bg-amber-200 text-amber-900 rounded font-bold">
                                      Duplicate
                                    </span>
                                  )}
                                </div>
                              </td>
                              <td className="py-3 px-4 text-slate-500 font-mono max-w-xs truncate" title={sampleVal}>
                                {sampleVal || <span className="italic text-slate-300">empty</span>}
                              </td>
                              <td className="py-3 px-4">
                                <select
                                  value={selectedValue}
                                  onChange={e => {
                                    const val = e.target.value;
                                    if (val === "ignore") {
                                      setMapping(prev => ({ ...prev, [col]: "ignore" }));
                                    } else {
                                      const fieldObj = formFields.find(f => f.id === Number(val));
                                      setMapping(prev => ({ ...prev, [col]: fieldObj || "ignore" }));
                                    }
                                  }}
                                  className="w-full text-xs bg-slate-50 border border-gray-300 rounded-lg py-1.5 px-2.5 focus:bg-white focus:outline-none focus:ring-1 focus:ring-blue-500 font-medium"
                                >
                                  <option value="ignore">[ -- Ignore Column -- ]</option>
                                  {formFields.map(f => (
                                    <option key={f.id} value={f.id}>
                                      {f.name} (ID: {f.id})
                                    </option>
                                  ))}
                                </select>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              </section>
            )}

            {/* STEP 4: Review Mapping */}
            {currentStep === 4 && (
              <section className="space-y-6 animate-in fade-in duration-150">
                <div className="bg-white rounded-2xl border border-gray-200 p-8 shadow-xs">
                  <div>
                    <h2 className="text-xl font-bold text-slate-900">Step 4: Review Active Field Mappings</h2>
                    <p className="text-sm text-slate-500 mt-1">
                      Verify configured column pairs before executing batch submission to Eloqua.
                    </p>
                  </div>

                  {/* 2-Column Table of Active Mappings */}
                  <div className="mt-6 border border-gray-200 rounded-xl overflow-hidden">
                    <table className="w-full text-left text-xs">
                      <thead className="bg-slate-50 text-slate-500 uppercase tracking-wider font-semibold border-b border-gray-200">
                        <tr>
                          <th className="py-3 px-4 w-1/2">Spreadsheet Column (Sample Value)</th>
                          <th className="py-3 px-4 w-1/2">Eloqua Target Field Name [ID]</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-100 text-slate-700">
                        {Object.entries(mapping)
                          .filter(([_, target]) => target !== "ignore")
                          .map(([col, target]) => {
                            const sample = uploadedData?.sample_row[col] ?? "";
                            const field = target as EloquaField;
                            return (
                              <tr key={col}>
                                <td className="py-2.5 px-4 font-semibold text-slate-800">
                                  {col} <span className="text-xs text-slate-400 font-normal">("{String(sample)}")</span>
                                </td>
                                <td className="py-2.5 px-4 text-blue-700 font-semibold">
                                  {field.name} <span className="text-xs text-slate-400 font-mono">[{field.id}]</span>
                                </td>
                              </tr>
                            );
                          })}
                      </tbody>
                    </table>
                  </div>

                  {/* Summary Bar */}
                  <div className="mt-6 p-4 bg-slate-50 border border-gray-200 rounded-xl flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs">
                    <div className="flex space-x-6 text-slate-600">
                      <div>
                        <span className="font-bold text-slate-900">{uploadedData?.total_rows}</span> total rows
                      </div>
                      <div>
                        <span className="font-bold text-emerald-600">{activeMappedCount}</span> columns mapped
                      </div>
                      <div>
                        <span className="font-bold text-slate-400">{ignoredCount}</span> ignored
                      </div>
                    </div>
                    <div className="text-slate-500">
                      Target Form: <span className="font-semibold text-slate-900">{selectedForm?.name}</span>
                    </div>
                  </div>
                </div>
              </section>
            )}

            {/* STEP 5: Process & Results */}
            {currentStep === 5 && (
              <section className="space-y-6 animate-in fade-in duration-150">
                <div className="bg-white rounded-2xl border border-gray-200 p-8 shadow-xs">
                  <div>
                    <h2 className="text-xl font-bold text-slate-900">Step 5: Processing Ingestion Batches</h2>
                    <p className="text-sm text-slate-500 mt-1">
                      Submitting records in batches of {config.batch_size} to Oracle Eloqua REST API 2.0 with {config.throttle_seconds}s pause between batches.
                    </p>
                  </div>

                  {/* Progress Bar */}
                  <div className="mt-6">
                    <div className="flex items-center justify-between text-xs font-semibold mb-2">
                      <span className="text-slate-700">
                        {isProcessing
                          ? `Submitting to Eloqua in batches of ${config.batch_size}...`
                          : circuitBreakerTriggered
                          ? 'Execution Halted by Circuit Breaker'
                          : 'Processing Complete'}
                      </span>
                      <span className="text-blue-600 font-mono">
                        {uploadedData?.total_rows
                          ? Math.round((processedRows / uploadedData.total_rows) * 100)
                          : 0}
                        %
                      </span>
                    </div>
                    <div className="w-full h-3 bg-gray-100 rounded-full overflow-hidden">
                      <div
                        className={`h-full transition-all duration-300 ${
                          circuitBreakerTriggered ? 'bg-rose-600' : 'bg-blue-600'
                        }`}
                        style={{
                          width: `${
                            uploadedData?.total_rows
                              ? Math.round((processedRows / uploadedData.total_rows) * 100)
                              : 0
                          }%`
                        }}
                      ></div>
                    </div>
                  </div>

                  {/* 4 Metric Cards */}
                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 mt-6">
                    <div className="p-4 bg-slate-50 border border-gray-200 rounded-xl">
                      <span className="text-xs text-slate-500 font-medium">Total Rows</span>
                      <p className="text-2xl font-bold text-slate-900 mt-1">{uploadedData?.total_rows || 0}</p>
                    </div>
                    <div className="p-4 bg-blue-50 border border-blue-200 rounded-xl">
                      <span className="text-xs text-blue-700 font-medium">Processed</span>
                      <p className="text-2xl font-bold text-blue-900 mt-1">{processedRows}</p>
                    </div>
                    <div className="p-4 bg-emerald-50 border border-emerald-200 rounded-xl">
                      <span className="text-xs text-emerald-700 font-medium">Successful</span>
                      <p className="text-2xl font-bold text-emerald-900 mt-1">{successfulRows}</p>
                    </div>
                    <div className="p-4 bg-rose-50 border border-rose-200 rounded-xl">
                      <span className="text-xs text-rose-700 font-medium">Failed</span>
                      <p className="text-2xl font-bold text-rose-900 mt-1">{failedRows}</p>
                    </div>
                  </div>

                  {/* Circuit Breaker Triggered Alert */}
                  {circuitBreakerTriggered && (
                    <div className="mt-6 p-4 bg-rose-50 border border-rose-300 rounded-xl">
                      <div className="flex items-start space-x-3">
                        <div className="w-8 h-8 rounded-full bg-rose-600 text-white flex items-center justify-center flex-shrink-0 font-bold">
                          !
                        </div>
                        <div>
                          <h4 className="text-sm font-bold text-rose-950">Circuit Breaker Triggered</h4>
                          <p className="text-xs text-rose-800 mt-1">{circuitBreakerReason}</p>
                        </div>
                      </div>
                    </div>
                  )}

                  {/* 100% Success Banner */}
                  {!isProcessing && !circuitBreakerTriggered && failedRows === 0 && processedRows > 0 && (
                    <div className="mt-6 p-6 bg-emerald-50 border border-emerald-200 rounded-xl text-center">
                      <div className="w-12 h-12 bg-emerald-600 text-white rounded-full flex items-center justify-center mx-auto mb-3 shadow-sm">
                        <Check className="w-6 h-6 stroke-[3]" />
                      </div>
                      <h3 className="text-base font-bold text-emerald-950">All Records Ingested Successfully!</h3>
                      <p className="text-xs text-emerald-700 mt-1">
                        All {successfulRows} records were transmitted and confirmed by Oracle Eloqua form #{selectedForm?.form_id}.
                      </p>
                    </div>
                  )}

                  {/* Failed Records Download Button */}
                  {failedRows > 0 && (
                    <div className="mt-6 text-center">
                      <button
                        onClick={handleDownloadFailedRecords}
                        className="inline-flex items-center space-x-2 px-5 py-2.5 bg-rose-600 hover:bg-rose-700 text-white text-xs font-bold rounded-xl shadow-md transition-colors"
                      >
                        <Download className="w-4 h-4" />
                        <span>Download Failed Records ({failedRows}) (.csv)</span>
                      </button>
                      <p className="text-[11px] text-slate-400 mt-2">
                        Contains all original spreadsheet columns plus appended Eloqua_HTTP_Status and Eloqua_Error_Details.
                      </p>
                    </div>
                  )}
                </div>
              </section>
            )}

          </main>

          {/* Sticky Bottom Action Bar */}
          <footer className="bg-white border-t border-gray-200 fixed bottom-0 left-0 right-0 z-40 py-3 shadow-lg">
            <div className="max-w-4xl mx-auto px-6 flex items-center justify-between">
              <div className="text-xs text-slate-500 font-medium">
                {currentStep === 1 && "Step 1 of 5: Select Eloqua destination form"}
                {currentStep === 2 && "Step 2 of 5: Upload spreadsheet (.xlsx or .csv)"}
                {currentStep === 3 && "Step 3 of 5: Map columns to Eloqua form fields"}
                {currentStep === 4 && "Step 4 of 5: Final review and confirmation"}
                {currentStep === 5 && (isProcessing ? "Ingestion dispatcher running..." : "Batch execution finished")}
              </div>

              <div className="flex space-x-3">
                {currentStep > 1 && currentStep < 5 && (
                  <button
                    onClick={() => setCurrentStep(prev => prev - 1)}
                    className="inline-flex items-center space-x-1 px-4 py-2 text-xs font-semibold text-slate-600 hover:text-slate-900 bg-gray-100 hover:bg-gray-200 rounded-xl transition-colors"
                  >
                    <ArrowLeft className="w-3.5 h-3.5" />
                    <span>Back</span>
                  </button>
                )}

                {currentStep === 1 && (
                  <button
                    onClick={() => setCurrentStep(2)}
                    disabled={!selectedForm}
                    className="inline-flex items-center space-x-1 px-5 py-2 text-xs font-bold text-white bg-blue-600 hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed rounded-xl shadow-sm transition-colors"
                  >
                    <span>Continue</span>
                    <ChevronRight className="w-3.5 h-3.5" />
                  </button>
                )}

                {currentStep === 2 && (
                  <button
                    onClick={() => setCurrentStep(3)}
                    disabled={!uploadedData}
                    className="inline-flex items-center space-x-1 px-5 py-2 text-xs font-bold text-white bg-blue-600 hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed rounded-xl shadow-sm transition-colors"
                  >
                    <span>Continue</span>
                    <ChevronRight className="w-3.5 h-3.5" />
                  </button>
                )}

                {currentStep === 3 && (
                  <button
                    onClick={() => setCurrentStep(4)}
                    disabled={!isEmailMapped()}
                    className="inline-flex items-center space-x-1 px-5 py-2 text-xs font-bold text-white bg-blue-600 hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed rounded-xl shadow-sm transition-colors"
                  >
                    <span>Review Mapping</span>
                    <ChevronRight className="w-3.5 h-3.5" />
                  </button>
                )}

                {currentStep === 4 && (
                  <button
                    onClick={() => {
                      setCurrentStep(5);
                      executeBatchProcessing();
                    }}
                    className="inline-flex items-center space-x-1.5 px-6 py-2 text-xs font-bold text-white bg-blue-600 hover:bg-blue-700 rounded-xl shadow-sm transition-colors"
                  >
                    <span>Submit to Eloqua Form</span>
                    <ChevronRight className="w-3.5 h-3.5" />
                  </button>
                )}

                {currentStep === 5 && !isProcessing && (
                  <button
                    onClick={resetAll}
                    className="inline-flex items-center space-x-1.5 px-5 py-2 text-xs font-bold text-slate-700 bg-gray-100 hover:bg-gray-200 rounded-xl transition-colors"
                  >
                    <RotateCcw className="w-3.5 h-3.5" />
                    <span>Upload Another Dataset</span>
                  </button>
                )}
              </div>
            </div>
          </footer>
        </>
      )}
    </div>
  );
}
