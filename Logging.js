/**
 * Sheet-based logging.
 *
 * Two sheets in the active spreadsheet, created on first use:
 *  - "log"       — general events: timestamp, name, severity, values (JSON), message
 *  - "email_log" — one row per email send attempt: timestamp, email_template,
 *                  recipient, quota_before, quota_after, status, payload (JSON)
 *
 * Logging never throws; failures are reported to console only so a broken
 * log sheet cannot stop a mail run.
 */

const LOG_SHEET_NAME = "log";
const EMAIL_LOG_SHEET_NAME = "email_log";

const LOG_HEADERS = ["timestamp", "name", "severity", "values", "message"];
const EMAIL_LOG_HEADERS = ["timestamp", "email_template", "recipient", "quota_before", "quota_after", "status", "payload"];

// Google Sheets cell limit is 50 000 characters; keep JSON safely under it.
const LOG_CELL_MAX_CHARS = 49000;
const LOG_TIMESTAMP_FORMAT = "dd.MM.yyyy HH:mm:ss";

/**
 * Remaining daily email quota for the effective user.
 * @return {number|null} Quota, or null when the value cannot be read.
 */
function getRemainingMailQuota_() {
  try {
    return MailApp.getRemainingDailyQuota();
  } catch (error) {
    console.warn(`Cannot read remaining mail quota: ${error.message}`);
    return null;
  }
}

/**
 * Logs the remaining daily quota to console and returns it.
 * @param {string} label Context label, e.g. "start of run".
 * @return {number|null}
 */
function logMailQuota_(label) {
  const quota = getRemainingMailQuota_();
  console.log(`Remaining daily email quota (${label}): ${quota === null ? "unknown" : quota}`);
  return quota;
}

/**
 * Quick manual check of the remaining daily email quota (run from the editor).
 */
function checkMailQuota() {
  logMailQuota_("manual check");
}

/**
 * JSON.stringify that survives Apps Script objects (Blobs become their
 * name, Dates their ISO string) and never exceeds the sheet cell limit.
 * @param {*} value
 * @return {string}
 */
function safeJsonForLog_(value) {
  if (value === undefined || value === null) return "";
  let json;
  try {
    json = JSON.stringify(value, (key, v) => {
      if (v && typeof v === "object" && typeof v.getName === "function" && typeof v.getBytes === "function") {
        return `<blob:${v.getName()}>`;
      }
      return v;
    });
  } catch (error) {
    json = String(value);
  }
  if (json === undefined) return "";
  return json.length > LOG_CELL_MAX_CHARS
    ? json.slice(0, LOG_CELL_MAX_CHARS - 3) + "..."
    : json;
}

/**
 * Returns the named sheet from the active spreadsheet, creating it with the
 * given header row when it does not exist yet.
 * @param {string} name
 * @param {string[]} headers
 * @return {Sheet}
 */
function getOrCreateLogSheet_(name, headers) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Appends a row to a log sheet with a real Date in the first column.
 * @param {string} sheetName
 * @param {string[]} headers
 * @param {Array} row Row values; index 0 must be the timestamp Date.
 */
function appendLogRow_(sheetName, headers, row) {
  const sheet = getOrCreateLogSheet_(sheetName, headers);
  sheet.appendRow(row);
  sheet.getRange(sheet.getLastRow(), 1).setNumberFormat(LOG_TIMESTAMP_FORMAT);
}

/**
 * Writes a general event to the "log" sheet. Never throws.
 * @param {string} name Event name, e.g. "mail_run_start".
 * @param {string} severity "INFO" | "WARNING" | "ERROR".
 * @param {object} [values] Structured values, stored as JSON.
 * @param {string} [message] Free-text message.
 */
function logEvent_(name, severity, values, message) {
  const text = message || "";
  const line = `[${severity}] ${name}${text ? ": " + text : ""}`;
  if (severity === "ERROR") console.error(line); else console.log(line);

  try {
    appendLogRow_(LOG_SHEET_NAME, LOG_HEADERS, [new Date(), name, severity, safeJsonForLog_(values), text]);
  } catch (error) {
    console.error(`Failed to write to "${LOG_SHEET_NAME}" sheet: ${error.message}`);
  }
}

/**
 * Sends an email through GmailApp and records the attempt in "email_log"
 * with the remaining quota before and after the send.
 * Rethrows the send error after logging it so callers keep their own handling.
 * @param {string} templateName Email template identifier (plan Email Topic).
 * @param {string} recipient
 * @param {string} subject
 * @param {string} body Plain-text body.
 * @param {object} options GmailApp.sendEmail options (from, cc, bcc, name, htmlBody, attachments, inlineImages).
 * @param {object} payload Data logged as JSON (template data, addressing, row reference).
 */
function sendEmailLogged_(templateName, recipient, subject, body, options, payload) {
  const quotaBefore = getRemainingMailQuota_();
  let status = "sent";
  let sendError = null;

  try {
    GmailApp.sendEmail(recipient, subject, body, options);
  } catch (error) {
    sendError = error;
    status = `ERROR: ${error.message}`;
  }

  const quotaAfter = getRemainingMailQuota_();
  console.log(`Email "${templateName}" to ${recipient}: ${status}. Remaining daily email quota: ${quotaAfter === null ? "unknown" : quotaAfter}`);

  try {
    appendLogRow_(EMAIL_LOG_SHEET_NAME, EMAIL_LOG_HEADERS, [
      new Date(),
      templateName,
      recipient,
      quotaBefore === null ? "" : quotaBefore,
      quotaAfter === null ? "" : quotaAfter,
      status,
      safeJsonForLog_(payload)
    ]);
  } catch (error) {
    console.error(`Failed to write to "${EMAIL_LOG_SHEET_NAME}" sheet: ${error.message}`);
  }

  if (sendError) throw sendError;
}
