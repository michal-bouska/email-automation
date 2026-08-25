// To learn how to use this script, refer to the documentation:
// https://developers.google.com/apps-script/samples/automations/mail-merge

/*
Copyright 2022 Martin Hawksey

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/
 
/**
 * Script accesses external spreadsheets via openById().
*/
 
/**
 * Change these to match the column names you are using for email 
 * recipient addresses and email sent column.
*/
const RECIPIENT_COL  = "Recipient";
const EMAIL_SENT_COL = "Email Sent";

function processEmailsMonitored() {
  withHealthcheck_(processEmails)();
}

function myFunction() {
  Logger.log("Toto je výchozí funkce.");
  processEmailsMonitored();
}
 
/** 
 * Creates the menu item "Mail Merge" for user to run scripts on drop-down.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
      .addItem('Send Emails', 'myFunction')
      .addToUi();
}

/**
 * Fill template string with data object
 * @see https://stackoverflow.com/a/378000/1027723
 * @param {string} template string containing {{}} markers which are replaced with data
 * @param {object} data object used to replace {{}} markers
 * @return {object} message replaced with data
*/
function fillInTemplateFromObject_(template, data) {
  // We have two templates one for plain text and the html body
  // Stringifing the object means we can do a global replace
  // Logger.log(`Try to fill template: Template = ${JSON.stringify(template, null, 2)}, Data = ${JSON.stringify(data, null, 2)}`);
  let template_string = JSON.stringify(template);

  // Token replacement
  template_string = template_string.replace(/{{[^{}]+}}/g, key => {
    return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
  });
  return  JSON.parse(template_string);
}



/**
 * Escape cell data to make JSON safe
 * @see https://stackoverflow.com/a/9204218/1027723
 * @param {string} str to escape JSON special characters from
 * @return {string} escaped string
*/
function escapeData_(str) {
  // Logger.log(`Escaping data: ${str}, Type: ${typeof str}`);
  str = String(str);
  return str
    .replace(/[\\]/g, '\\\\')
    .replace(/[\"]/g, '\\\"')
    .replace(/[\/]/g, '\\/')
    .replace(/[\b]/g, '\\b')
    .replace(/[\f]/g, '\\f')
    .replace(/[\n]/g, '\\n')
    .replace(/[\r]/g, '\\r')
    .replace(/[\t]/g, '\\t');
}

/**
 * Module-level cache for Gmail drafts and processed templates.
 * Fetches drafts lazily on first access and caches processed templates by subject line.
 */
const DRAFT_CACHE = (() => {
  let drafts = null;
  const templates = {};
  const notFound = {};
  const htmlTemplates = {};

  return {
    getDrafts() {
      if (drafts === null) {
        console.log("Gmail API call: GmailApp.getDrafts() - fetching all drafts.");
        drafts = GmailApp.getDrafts();
        console.log(`Gmail API call finished: fetched ${drafts.length} drafts.`);
      } else {
        console.log("Using cached Gmail drafts (no API call).");
      }
      return drafts;
    },
    getTemplate(subjectLine) {
      if (templates[subjectLine]) return templates[subjectLine];
      if (notFound[subjectLine]) return null;
      return undefined;
    },
    setTemplate(subjectLine, template) {
      templates[subjectLine] = template;
    },
    setNotFound(subjectLine) {
      notFound[subjectLine] = true;
    },
    getHtmlTemplate(filename) {
      return htmlTemplates[filename] || undefined;
    },
    setHtmlTemplate(filename, template) {
      htmlTemplates[filename] = template;
    }
  };
})();

function getGmailTemplateFromDrafts_(subject_line){
  const cached = DRAFT_CACHE.getTemplate(subject_line);
  if (cached) {
    console.log(`Draft template for subject "${subject_line}" served from cache.`);
    return cached;
  }
  if (cached === null) throw new Error("Oops - can't find Gmail draft");

  try {
    console.log(`Looking up Gmail draft with subject "${subject_line}".`);
    const drafts = DRAFT_CACHE.getDrafts();
    const draft = drafts.filter(subjectFilter_(subject_line))[0];

    if (!draft) {
      console.log(`Gmail draft with subject "${subject_line}" not found.`);
      DRAFT_CACHE.setNotFound(subject_line);
      throw new Error("Oops - can't find Gmail draft");
    }

    console.log(`Gmail API call: loading draft message and attachments for subject "${subject_line}".`);
    const msg = draft.getMessage();

    // Handles inline images and attachments so they can be included in the merge
    // Based on https://stackoverflow.com/a/65813881/1027723
    const allInlineImages = draft.getMessage().getAttachments({includeInlineImages: true, includeAttachments: false});
    const attachments = draft.getMessage().getAttachments({includeInlineImages: false});
    const htmlBody = msg.getBody();

    const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj), {});
    const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
    const matches = [...htmlBody.matchAll(imgexp)];

    const inlineImagesObj = {};
    matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

    const template = {
      message: {subject: subject_line, text: msg.getPlainBody(), html: htmlBody},
      attachments: attachments,
      inlineImages: inlineImagesObj
    };

    console.log(`Draft template for subject "${subject_line}" processed and cached.`);
    DRAFT_CACHE.setTemplate(subject_line, template);
    return template;
  } catch(e) {
    if (e.message === "Oops - can't find Gmail draft") throw e;
    DRAFT_CACHE.setNotFound(subject_line);
    throw new Error("Oops - can't find Gmail draft");
  }

  function subjectFilter_(subject_line){
    return function(element) {
      if (element.getMessage().getSubject() === subject_line) {
        return element;
      }
    }
  }
}

/**
 * Loads an HTML email template from a Google Drive folder.
 * Folder ID is configured as Script Property "HTML_TEMPLATES_FOLDER_ID".
 * @param {string} filename The HTML file name (e.g., "invoice.html").
 * @param {string} subjectLine The subject line for the email.
 * @return {object} Template object matching getGmailTemplateFromDrafts_ structure.
 */
function getHtmlTemplateFromDrive_(filename, subjectLine) {
  const folderId = PropertiesService.getScriptProperties().getProperty('HTML_TEMPLATES_FOLDER_ID');
  if (!folderId) {
    throw new Error("Script Property 'HTML_TEMPLATES_FOLDER_ID' is not configured.");
  }

  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByName(filename);

  if (!files.hasNext()) {
    throw new Error(`HTML template file "${filename}" not found in Drive folder.`);
  }

  const file = files.next();
  const htmlContent = file.getBlob().getDataAsString();

  const plainText = htmlContent
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<\/li>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

  return {
    message: { subject: subjectLine, text: plainText, html: htmlContent },
    attachments: [],
    inlineImages: {}
  };
}

/**
 * Unified template loader. Selects source based on plan's templateSource.
 * @param {object} plan The parsed plan row object.
 * @return {object} Template object with { message, attachments, inlineImages }.
 */
function getEmailTemplate_(plan) {
  const source = (plan.templateSource || "").trim().toLowerCase();

  if (source === "" || source === "draft") {
    return getGmailTemplateFromDrafts_(plan.emailTopic);
  }

  const filename = plan.templateSource.trim();
  const cached = DRAFT_CACHE.getHtmlTemplate(filename);
  if (cached) return cached;

  const template = getHtmlTemplateFromDrive_(filename, plan.emailTopic);
  DRAFT_CACHE.setHtmlTemplate(filename, template);
  return template;
}

/**
 * Generates a QR code as a blob (image data) using QuickChart.
 * @param {object} qrCodeObj The QR code object with all necessary fields.
 * @return {Blob} The QR code image as a blob, or null if data is missing.
 */
function generateQrCodeBlob(qrCodeObj) {
  if (!qrCodeObj) {
    console.error("QR code data object is missing.");
    return null;
  }

  const dataString = generatePaymentQrData(
    qrCodeObj.accountNumber,
    qrCodeObj.bankCode,
    qrCodeObj.currency,
    qrCodeObj.amount,
    qrCodeObj.variableSymbol,
    qrCodeObj.message
  );

  const chartUrl = `https://quickchart.io/qr?text=${encodeURIComponent(dataString)}&size=${qrCodeObj.size}`;

  try {
    const response = UrlFetchApp.fetch(chartUrl);
    if (response.getResponseCode() === 200) {
      return response.getBlob();
    } else {
      console.error(`Failed to fetch QR code. HTTP response code ${response.getResponseCode()}.`);
      return null;
    }
  } catch (e) {
    console.error(`Error generating QR code: ${e.message}`);
    return null;
  }
}

/**
 * Generates data string for payment QR code.
 * @param {string} accountNumber The account number of the recipient.
 * @param {string} bankCode The bank code of the recipient.
 * @param {string} currency The currency of the payment.
 * @param {number} amount The amount of the payment.
 * @param {string} variableSymbol The variable symbol for the transaction.
 * @param {string} message An optional message for the payment.
 * @return {string} The formatted payment data string.
 */
function generatePaymentQrData(accountNumber, bankCode, currency, amount, variableSymbol, message) {
  let iban = convertToIBAN(accountNumber, bankCode);
  let dataString = `SPD*1.0*ACC:${iban}*AM:${amount.toFixed(2)}*CC:${currency}`;

  if (variableSymbol) {
    dataString += `*VS:${variableSymbol}`;
  }

  if (message) {
    dataString += `*MSG:${message}`;
  }

  return dataString;
}

/**
 * Consolidates data for QR code generation.
 * @param {object} qrCodeObj The QR code object with all necessary fields.
 * @param {Array} realizationData The realization data (rows from the sheet).
 * @param {Array} realizationHeaders The headers for the realization data.
 * @return {object} Consolidated QR code data with resolved variable symbol.
 */
function consolidateQrCodeData(qrCodeObj, realizationData, realizationHeaders) {
  if (!qrCodeObj) {
    throw new Error("QR code object is required.");
  }

  if (qrCodeObj.variableSymbol && qrCodeObj.variableSymbolColumn) {
    throw new Error("Only one of variableSymbol or variableSymbolColumn should be non-empty.");
  }

  let resolvedVariableSymbol = qrCodeObj.variableSymbol;

  if (!resolvedVariableSymbol && qrCodeObj.variableSymbolColumn) {
    const columnIdx = realizationHeaders.indexOf(qrCodeObj.variableSymbolColumn);
    if (columnIdx === -1) {
      throw new Error(`Column ${qrCodeObj.variableSymbolColumn} not found in realization data headers.`);
    }

    // Assuming the first row of realizationData is the one to use for variable symbol
    resolvedVariableSymbol = realizationData[0][columnIdx];

    if (resolvedVariableSymbol == null) {
      throw new Error(`No value found in realization data for column ${qrCodeObj.variableSymbolColumn}.`);
    }
  }

  return {
    ...qrCodeObj,
    variableSymbol: resolvedVariableSymbol
  };
}


/**
 * Module-level plan data loader. Only loads the plan sheet from the active spreadsheet.
 */
const PLAN_DATA = (() => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const planSheet = spreadsheet.getSheetByName("plan");
  return planSheet ? planSheet.getDataRange().getValues() : [];
})();

/**
 * Cache for opened external spreadsheets, keyed by Document ID.
 */
const SPREADSHEET_CACHE = (() => {
  const cache = {};
  return {
    get(documentId) {
      if (!cache[documentId]) {
        cache[documentId] = SpreadsheetApp.openById(documentId);
      }
      return cache[documentId];
    }
  };
})();

/**
 * Parses QR code sheet data into objects mapped by Email Topic.
 * @param {any[][]} qrCodeData Raw 2D array from the QRCodes sheet (including header row).
 * @return {object} Map of emailTopic -> Array of qrCodeObj.
 */
function parseQrCodeData(qrCodeData) {
  if (!qrCodeData || qrCodeData.length < 2) {
    return {};
  }

  const qrCodeHeaders = qrCodeData[0];
  return qrCodeData.slice(1).reduce((map, row) => {
    const qrCodeObj = {
      emailTopic: row[qrCodeHeaders.indexOf("EmailTopic")],
      imageName: row[qrCodeHeaders.indexOf("ImageName")],
      accountNumber: row[qrCodeHeaders.indexOf("AccountNumber")],
      bankCode: row[qrCodeHeaders.indexOf("BankCode")],
      currency: row[qrCodeHeaders.indexOf("Currency")],
      amount: row[qrCodeHeaders.indexOf("Amount")],
      variableSymbol: row[qrCodeHeaders.indexOf("VariableSymbol")],
      variableSymbolColumn: row[qrCodeHeaders.indexOf("VariableSymbolColumn")],
      message: row[qrCodeHeaders.indexOf("Message")],
      size: row[qrCodeHeaders.indexOf("Size")]
    };

    if (qrCodeObj.emailTopic) {
      if (!map[qrCodeObj.emailTopic]) {
        map[qrCodeObj.emailTopic] = [];
      }
      map[qrCodeObj.emailTopic].push(qrCodeObj);
    }

    return map;
  }, {});
}

/**
 * Loads realization data and QR code data from an external spreadsheet.
 * @param {string} documentId The Google Spreadsheet ID.
 * @param {string} sheetName The sheet tab name containing realization data.
 * @return {object} { headers, data, qrCodes, sheet }
 */
function loadRealizationDocument(documentId, sheetName) {
  const spreadsheet = SPREADSHEET_CACHE.get(documentId);

  const realizationSheet = spreadsheet.getSheetByName(sheetName);
  if (!realizationSheet) {
    throw new Error(`Sheet "${sheetName}" not found in document ${documentId}.`);
  }

  const realizationRaw = realizationSheet.getDataRange().getValues();
  const headers = realizationRaw[0] || [];
  const data = realizationRaw.slice(1);

  const qrCodeSheet = spreadsheet.getSheetByName("QRCodes");
  let qrCodes = {};
  if (qrCodeSheet) {
    qrCodes = parseQrCodeData(qrCodeSheet.getDataRange().getValues());
  }

  return { headers, data, qrCodes, sheet: realizationSheet };
}

/**
 * Transforms plan data into objects with validation.
 * Plan sheet columns: Email Topic, Column Condition To Send, Column Sent, Document ID, Sheet Name.
 * @returns {Array<Object>} Array of parsed plan objects.
 */
function parsePlanData() {
  const planData = PLAN_DATA;
  if (!planData || planData.length < 2) {
    throw new Error("Plan data is missing or insufficient rows.");
  }

  const headers = planData[0];
  const emailTopicIdx = headers.indexOf("Email Topic");
  const conditionColumnIdx = headers.indexOf("Column Condition To Send");
  const sentDateColumnIdx = headers.indexOf("Column Sent");
  const documentIdIdx = headers.indexOf("Document ID");
  const sheetNameIdx = headers.indexOf("Sheet Name");
  const templateSourceIdx = headers.indexOf("Template Source");
  const issueNameIdx = headers.indexOf("Issue Name");
  const issueColumnIdx = headers.indexOf("Issue Column");

  if (emailTopicIdx === -1 || conditionColumnIdx === -1 || sentDateColumnIdx === -1) {
    throw new Error("Required headers are missing in the plan data.");
  }
  if (documentIdIdx === -1 || sheetNameIdx === -1) {
    throw new Error("Document ID and Sheet Name columns are required in the plan data.");
  }

  return planData.slice(1)
    .filter(row => row[emailTopicIdx])
    .map(row => {
      const documentId = row[documentIdIdx];
      const sheetName = row[sheetNameIdx];

      if (!documentId || !sheetName) {
        throw new Error(
          `Plan row for topic "${row[emailTopicIdx]}" is missing Document ID or Sheet Name.`
        );
      }

      return {
        emailTopic: row[emailTopicIdx],
        conditionColumn: row[conditionColumnIdx],
        sentColumn: row[sentDateColumnIdx],
        documentId: String(documentId),
        sheetName: String(sheetName),
        templateSource: templateSourceIdx !== -1 ? String(row[templateSourceIdx] || "") : "",
        issueName: issueNameIdx !== -1 ? String(row[issueNameIdx] || "").trim() : "",
        issueColumn: issueColumnIdx !== -1 ? String(row[issueColumnIdx] || "").trim() : ""
      };
    });
}

/**
 * Inserts QR codes into email content by replacing placeholders.
 * @param {string} htmlBody The email HTML body.
 * @param {object} inlineImages Object containing inline images to attach.
 * @param {object} qrCodeObj The QR code object with all necessary fields.
 * @return {string} Updated HTML body with QR code placeholders replaced.
 */
function insertQrCodesIntoEmail(htmlBody, inlineImages, qrCodeObj) {
  Logger.log("Start insertQrCodesIntoEmail")
  const qrCodeBlob = generateQrCodeBlob(qrCodeObj);
  if (qrCodeBlob) {
    inlineImages[qrCodeObj.imageName] = qrCodeBlob;
    htmlBody = htmlBody.replace(`{{${qrCodeObj.imageName}}}`, `<img data-surl="cid:${qrCodeObj.imageName}" src="cid:${qrCodeObj.imageName}" alt="QR Code">`);
  }
  return htmlBody;
}

function processEmails() {
  const planData = parsePlanData();
  const realizationCache = {};

  // Collects all per-campaign and per-recipient errors; reported to the
  // HEALTHCHECKS_PING_URL_WARNING check at the end of the run.
  const runErrors = [];
  const reportRunError = (message) => {
    console.error(message);
    runErrors.push(message);
  };

  planData.forEach(plan => {
    const cacheKey = `${plan.documentId}|${plan.sheetName}`;

    let realizationDoc;
    if (realizationCache[cacheKey]) {
      realizationDoc = realizationCache[cacheKey];
    } else {
      try {
        realizationDoc = loadRealizationDocument(plan.documentId, plan.sheetName);
        realizationCache[cacheKey] = realizationDoc;
      } catch (error) {
        reportRunError(`Failed to load realization document for topic "${plan.emailTopic}": ${error.message}`);
        return;
      }
    }

    const { headers: realizationHeaders, data: realizationData, qrCodes, sheet: realizationSheet } = realizationDoc;

    const recipientIdx = realizationHeaders.indexOf(RECIPIENT_COL);
    if (recipientIdx === -1) {
      reportRunError(`Recipient column (${RECIPIENT_COL}) not found in document ${plan.documentId}, sheet ${plan.sheetName}.`);
      return;
    }

    const conditionColIdx = realizationHeaders.indexOf(plan.conditionColumn);
    const sentColIdx = realizationHeaders.indexOf(plan.sentColumn);

    if (conditionColIdx === -1 || sentColIdx === -1) {
      reportRunError(`Columns "${plan.conditionColumn}" or "${plan.sentColumn}" not found in document ${plan.documentId}, sheet ${plan.sheetName}.`);
      return;
    }

    realizationData.forEach((recipientRow, recipientIndex) => {
      const recipient = recipientRow[recipientIdx];
      const conditionValue = recipientRow[conditionColIdx];
      const sentValue = recipientRow[sentColIdx];

      if (conditionValue !== 1 || sentValue !== "") {
        return;
      }

      console.log(`Preparing to send email for topic: ${plan.emailTopic} to recipient: ${recipient} at row ${recipientIndex + 2}.`);

      const writeSentStatus = (status) => {
        realizationSheet.getRange(recipientIndex + 2, sentColIdx + 1).setValue(status);
      };

      let emailTemplate, dataMapping, msgObj, attachments, inlineImages;

      try {
        emailTemplate = getEmailTemplate_(plan);
      } catch (error) {
        reportRunError(`Failed to get template for ${recipient}: ${error.message}`);
        writeSentStatus(`${error.message} at ${new Date().toISOString()}`);
        return;
      }

      try {
        dataMapping = realizationHeaders.reduce((map, header, index) => {
          map[header] = recipientRow[index];
          return map;
        }, {});
      } catch (error) {
        reportRunError(`Failed to parse realization for ${recipient}: ${error.message}`);
        writeSentStatus(`${error.message} at ${new Date().toISOString()}`);
        return;
      }

      try {
        if (qrCodes[plan.emailTopic]) {
          qrCodes[plan.emailTopic].forEach(qrCodeObj => {
            consolidateQrCodeData(qrCodeObj, [recipientRow], realizationHeaders);
            dataMapping[qrCodeObj.imageName] = `<img src="cid:${qrCodeObj.imageName}" alt="QR Code">`;
          });
        }
      } catch (error) {
        reportRunError(`Failed to generate QR code for ${recipient}: ${error.message}`);
        writeSentStatus(`${error.message} at ${new Date().toISOString()}`);
        return;
      }

      try {
        msgObj = fillInTemplateFromObject_(emailTemplate.message, {
          ...dataMapping,
          recipient,
          conditionValue,
          sentValue
        });
      } catch (error) {
        reportRunError(`Failed to fill template for ${recipient}: ${error.message}`);
        writeSentStatus(`${error.message} at ${new Date().toISOString()}`);
        return;
      }

      try {
        attachments = emailTemplate.attachments || [];
        inlineImages = emailTemplate.inlineImages || {};

        if (qrCodes[plan.emailTopic]) {
          qrCodes[plan.emailTopic].forEach(qrCodeObj => {
            const consolidated = consolidateQrCodeData(qrCodeObj, [recipientRow], realizationHeaders);
            msgObj.html = insertQrCodesIntoEmail(msgObj.html, inlineImages, consolidated);
          });
        }
      } catch (error) {
        reportRunError(`Failed to inline images for ${recipient}: ${error.message}`);
        writeSentStatus(`${error.message} at ${new Date().toISOString()}`);
        return;
      }

      try {
        GmailApp.sendEmail(recipient, msgObj.subject, msgObj.text, {
          htmlBody: msgObj.html,
          attachments: attachments,
          inlineImages: inlineImages
        });

        const range = realizationSheet.getRange(recipientIndex + 2, sentColIdx + 1);
        range.setValue(new Date().toISOString());
        range.setNumberFormat("dd.MM.yyyy HH:mm:ss");
      } catch (error) {
        reportRunError(`Failed to send email to ${recipient}: ${error.message}`);
        writeSentStatus(`${error.message} at ${new Date().toISOString()}`);
      }
    });
  });

  // Warning check: fails when any single campaign or email failed, even
  // though the run itself finished. Success ping keeps the check alive.
  if (runErrors.length > 0) {
    pingHealthcheck_('/fail', `${runErrors.length} error(s) during run:\n\n${runErrors.join('\n')}`, 'HEALTHCHECKS_PING_URL_WARNING');
  } else {
    pingHealthcheck_('', null, 'HEALTHCHECKS_PING_URL_WARNING');
  }
}


 
/**
 * Sends emails from sheet data.
 * @param {string} subjectLine (optional) for the email draft message
 * @param {Sheet} sheet to read data from
*/
function sendEmails(subjectLine, sheet=SpreadsheetApp.getActiveSheet()) {
  // option to skip browser prompt if you want to use this code in other projects
  
  if (!subjectLine){
    subjectLine = Browser.inputBox("Mail Merge", 
                                      "Type or copy/paste the subject line of the Gmail " +
                                      "draft message you would like to mail merge with:",
                                      Browser.Buttons.OK_CANCEL);
                                      
    if (subjectLine === "cancel" || subjectLine == ""){ 
    // If no subject line, finishes up
    return;
    }
  }
  
  // Gets the draft Gmail message to use as a template
  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
  
  // Gets the data from the passed sheet
  const dataRange = sheet.getDataRange();
  // Fetches displayed values for each row in the Range HT Andrew Roberts 
  // https://mashe.hawksey.info/2020/04/a-bulk-email-mail-merge-with-gmail-and-google-sheets-solution-evolution-using-v8/#comment-187490
  // @see https://developers.google.com/apps-script/reference/spreadsheet/range#getdisplayvalues
  const data = dataRange.getDisplayValues();

  // Assumes row 1 contains our column headings
  const heads = data.shift(); 
  
  // Gets the index of the column named 'Email Status' (Assumes header names are unique)
  // @see http://ramblings.mcpher.com/Home/excelquirks/gooscript/arrayfunctions
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);
  
  // Converts 2d array into an object array
  // See https://stackoverflow.com/a/22917499/1027723
  // For a pretty version, see https://mashe.hawksey.info/?p=17869/#comment-184945
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  // Creates an array to record sent emails
  const out = [];

  // Loops through all the rows of data
  obj.forEach(function(row, rowIdx){
    // Only sends emails if email_sent cell is blank and not hidden by a filter
    if (row[EMAIL_SENT_COL] == ''){
      try {
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);

        // See https://developers.google.com/apps-script/reference/gmail/gmail-app#sendEmail(String,String,String,Object)
        // If you need to send emails with unicode/emoji characters change GmailApp for MailApp
        // Uncomment advanced parameters as needed (see docs for limitations)
        GmailApp.sendEmail(row[RECIPIENT_COL], msgObj.subject, msgObj.text, {
          htmlBody: msgObj.html,
          // bcc: 'a.bcc@email.com',
          // cc: 'a.cc@email.com',
          // from: 'an.alias@email.com',
          // name: 'name of the sender',
          // replyTo: 'a.reply@email.com',
          // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
          attachments: emailTemplate.attachments,
          inlineImages: emailTemplate.inlineImages
        });
        // Edits cell to record email sent date
        out.push([new Date()]);
      } catch(e) {
        // modify cell to record error
        out.push([e.message]);
      }
    } else {
      out.push([row[EMAIL_SENT_COL]]);
    }
  });
  
  // Updates the sheet with new data
  sheet.getRange(2, emailSentColIdx+1, out.length).setValues(out);
  
  /**
   * Get a Gmail draft message by matching the subject line.
   * @param {string} subject_line to search for draft message
   * @return {object} containing the subject, plain and html message body and attachments
  */

}
