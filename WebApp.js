/**
 * Web UI for issuing items ("výdej") to recipients from realization documents.
 *
 * Issue points are defined in the plan sheet: every row with non-empty
 * "Issue Name" and "Issue Column" appears in the web form. Issuing writes
 * a timestamp into the issue column of the campaign's realization sheet.
 *
 * Deployed as a web app (Deploy > New deployment > Web app).
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Výdej')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Lists issue point names defined in the plan.
 * @return {string[]} Issue names.
 */
function getIssuePoints() {
  const seen = {};
  return parsePlanData()
    .filter(plan => plan.issueName && plan.issueColumn)
    .map(plan => plan.issueName)
    .filter(name => (seen[name] ? false : (seen[name] = true)));
}

/**
 * Finds the plan row for an issue point.
 * @param {string} issueName The issue point name.
 * @return {object} The plan object.
 */
function findIssuePlan_(issueName) {
  const plan = parsePlanData().find(p => p.issueName === issueName && p.issueColumn);
  if (!plan) {
    throw new Error(`Výdejové místo "${issueName}" nebylo nalezeno v plánu.`);
  }
  return plan;
}

/**
 * Loads the list of recipients for an issue point.
 * @param {string} issueName The issue point name.
 * @return {object} { issueName, issueColumn, items: [{rowNumber, name, recipient, issued}] }
 */
function getIssueList(issueName) {
  const plan = findIssuePlan_(issueName);

  // Lightweight load: unlike loadRealizationDocument, skips the QRCodes sheet.
  const sheet = SPREADSHEET_CACHE.get(plan.documentId).getSheetByName(plan.sheetName);
  if (!sheet) {
    throw new Error(`List "${plan.sheetName}" nebyl nalezen.`);
  }
  const values = sheet.getDataRange().getValues();
  const headers = values[0] || [];
  const data = values.slice(1);

  const issueColIdx = headers.indexOf(plan.issueColumn);
  if (issueColIdx === -1) {
    throw new Error(`Sloupec "${plan.issueColumn}" nebyl nalezen v listu "${plan.sheetName}".`);
  }

  const recipientIdx = headers.indexOf(RECIPIENT_COL);
  const firstNameIdx = headers.indexOf("First name");
  const lastNameIdx = headers.indexOf("Last name");

  const items = data
    .map((row, i) => ({
      rowNumber: i + 2,
      name: [
        firstNameIdx !== -1 ? row[firstNameIdx] : "",
        lastNameIdx !== -1 ? row[lastNameIdx] : ""
      ].join(" ").trim(),
      recipient: recipientIdx !== -1 ? String(row[recipientIdx] || "") : "",
      issued: formatIssuedValue_(row[issueColIdx])
    }))
    .filter(item => item.name || item.recipient);

  return { issueName: plan.issueName, issueColumn: plan.issueColumn, items: items };
}

/**
 * Marks a row as issued (or clears the mark) in the realization sheet.
 * @param {string} issueName The issue point name.
 * @param {number} rowNumber 1-based sheet row number.
 * @param {string} recipientCheck Expected Recipient value, guards against stale data.
 * @param {boolean} issued True to mark issued, false to clear.
 * @param {string} operator Free-text issue desk name, stored next to the timestamp.
 * @return {object} { rowNumber, issued } with the stored value.
 */
function markIssued(issueName, rowNumber, recipientCheck, issued, operator) {
  const plan = findIssuePlan_(issueName);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SPREADSHEET_CACHE.get(plan.documentId).getSheetByName(plan.sheetName);
    if (!sheet) {
      throw new Error(`List "${plan.sheetName}" nebyl nalezen.`);
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const issueColIdx = headers.indexOf(plan.issueColumn);
    if (issueColIdx === -1) {
      throw new Error(`Sloupec "${plan.issueColumn}" nebyl nalezen v listu "${plan.sheetName}".`);
    }

    const recipientIdx = headers.indexOf(RECIPIENT_COL);
    if (recipientIdx !== -1 && recipientCheck) {
      const actual = String(sheet.getRange(rowNumber, recipientIdx + 1).getValue() || "");
      if (actual !== recipientCheck) {
        throw new Error("Řádky v tabulce se mezitím změnily, obnovte si prosím seznam.");
      }
    }

    // Double-issue guard: never overwrite an existing issue record. Returns
    // the stored value with a conflict flag so the UI can inform the operator.
    const currentValue = formatIssuedValue_(sheet.getRange(rowNumber, issueColIdx + 1).getValue());
    if (issued && currentValue) {
      console.log(`Issue "${issueName}": row ${rowNumber} (${recipientCheck}) already issued: "${currentValue}".`);
      return { rowNumber: rowNumber, issued: currentValue, conflict: true };
    }

    const desk = String(operator || "").trim();
    const value = issued
      ? Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss")
        + (desk ? ` – ${desk}` : "")
      : "";
    sheet.getRange(rowNumber, issueColIdx + 1).setValue(value);

    console.log(`Issue "${issueName}": row ${rowNumber} (${recipientCheck}) set to "${value}".`);
    return { rowNumber: rowNumber, issued: value };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Formats a raw issue-column cell value for display.
 * @param {*} value Raw cell value (Date, string or empty).
 * @return {string} Display value; empty string when not issued.
 */
function formatIssuedValue_(value) {
  if (value === null || value === undefined || value === "") return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss");
  }
  return String(value);
}
