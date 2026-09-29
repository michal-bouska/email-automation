/**
 * Script Dashboard: job registry, last-run tracking and the "Script Dashboard"
 * sheet showing trigger state and the last run of every job.
 *
 * Last-run records live in Script Properties (JOB_RUN_<job>) so they survive
 * between executions; the sheet is a rendered view refreshed from the menu
 * and automatically after every tracked job run.
 */

const DASHBOARD_SHEET_NAME = "Script Dashboard";
const DASHBOARD_TIMESTAMP_FORMAT = "dd.MM.yyyy HH:mm:ss";
const JOB_RUN_PROPERTY_PREFIX = "JOB_RUN_";

/**
 * Registry of scheduled jobs. Single source of truth for trigger handlers
 * and intervals used by the trigger create/delete functions and the dashboard.
 */
const JOBS = {
  mail: { label: "Mail merge", handler: "processEmailsMonitored", intervalMinutes: 60 },
  fio:  { label: "Fio sync",   handler: "runUpdateMonitored",     intervalMinutes: 5 }
};

/**
 * Wraps a job so its start, end, duration, status and result are recorded
 * in Script Properties and the dashboard sheet is refreshed afterwards.
 * Tracking never throws; the job's own error is rethrown untouched.
 * @param {string} jobName Key in JOBS.
 * @param {Function} fn Job body; its return value is stored as run details.
 * @return {Function} Wrapped function.
 */
function withJobTracking_(jobName, fn) {
  return function(...args) {
    const startedAt = new Date();
    const record = { startedAt: startedAt.toISOString(), finishedAt: null, status: "running", details: null, error: null };
    saveJobRun_(jobName, record);

    try {
      const result = fn.apply(this, args);
      record.status = "ok";
      record.details = result === undefined ? null : result;
      return result;
    } catch (error) {
      record.status = "error";
      record.error = error.message;
      throw error;
    } finally {
      record.finishedAt = new Date().toISOString();
      saveJobRun_(jobName, record);
      try {
        refreshDashboard();
      } catch (error) {
        console.error(`Dashboard refresh failed (non-blocking): ${error.message}`);
      }
    }
  };
}

/**
 * Persists a job run record. Never throws.
 * @param {string} jobName
 * @param {object} record
 */
function saveJobRun_(jobName, record) {
  try {
    PropertiesService.getScriptProperties().setProperty(JOB_RUN_PROPERTY_PREFIX + jobName, JSON.stringify(record));
  } catch (error) {
    console.error(`Failed to save run record for job "${jobName}": ${error.message}`);
  }
}

/**
 * Loads the last run record of a job.
 * @param {string} jobName
 * @return {object|null}
 */
function getJobRun_(jobName) {
  const raw = PropertiesService.getScriptProperties().getProperty(JOB_RUN_PROPERTY_PREFIX + jobName);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (error) {
    console.warn(`Corrupt run record for job "${jobName}": ${error.message}`);
    return null;
  }
}

/**
 * Describes the installed triggers of a job.
 * @param {object} job Entry from JOBS.
 * @return {string} e.g. "active (1)" or "none".
 */
function describeTriggerState_(job) {
  const count = ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === job.handler)
    .length;
  return count > 0 ? `active (${count})` : "none";
}

/**
 * Human-readable one-line summary of a job's result details.
 * @param {*} details
 * @return {string}
 */
function formatJobDetails_(details) {
  if (details === null || details === undefined) return "";
  if (typeof details !== "object") return String(details);
  return Object.keys(details).map(key => `${key}: ${details[key]}`).join(", ");
}

/**
 * Rebuilds the "Script Dashboard" sheet. Contents are replaced, formats and
 * column widths are left alone. Safe to run from the menu at any time.
 */
function refreshDashboard() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(DASHBOARD_SHEET_NAME);
  }

  const now = new Date();
  const quota = getRemainingMailQuota_();
  const headers = ["Job", "Trigger", "Handler", "Interval", "Last start", "Last end", "Duration (s)", "Status", "Details", "Error"];

  const rows = [];
  const dateCells = []; // [row (1-based), col (1-based), Date]

  rows.push(["Script Dashboard", "updated", null]);
  dateCells.push([1, 3, now]);
  rows.push([]);
  rows.push(headers);

  Object.keys(JOBS).forEach(jobName => {
    const job = JOBS[jobName];
    const run = getJobRun_(jobName);
    const rowNumber = rows.length + 1;

    const startedAt = run && run.startedAt ? new Date(run.startedAt) : null;
    const finishedAt = run && run.finishedAt ? new Date(run.finishedAt) : null;
    const duration = startedAt && finishedAt ? Math.round((finishedAt - startedAt) / 1000) : "";

    rows.push([
      job.label,
      describeTriggerState_(job),
      job.handler,
      `every ${job.intervalMinutes} min`,
      null,
      null,
      duration,
      run ? run.status : "never run",
      run ? formatJobDetails_(run.details) : "",
      run && run.error ? run.error : ""
    ]);
    if (startedAt) dateCells.push([rowNumber, 5, startedAt]);
    if (finishedAt) dateCells.push([rowNumber, 6, finishedAt]);
  });

  rows.push([]);
  rows.push(["Remaining daily email quota", quota === null ? "unknown" : quota]);

  const width = Math.max(...rows.map(r => r.length));
  const grid = rows.map(r => { const row = r.slice(); while (row.length < width) row.push(""); return row.map(v => v === null ? "" : v); });

  sheet.clearContents();
  sheet.getRange(1, 1, grid.length, width).setValues(grid);
  dateCells.forEach(([row, col, date]) => setDateKeepingFormat_(sheet.getRange(row, col), date, DASHBOARD_TIMESTAMP_FORMAT));
  sheet.getRange(3, 1, 1, headers.length).setFontWeight("bold");
  sheet.getRange(1, 1).setFontWeight("bold");
  sheet.setFrozenRows(3);

  console.log(`Dashboard refreshed at ${now.toISOString()}.`);
}
