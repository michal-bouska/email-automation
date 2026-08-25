const MAIL_CONFIG = loadConfig([["RECIPIENT_COL", "Recipient"], ["EMAIL_PLAN_SHEET", "plan"], ["EMAIL_LOG_SHEET", "realizace"], ["SENDER_EMAIL", null], ["CC_EMAIL", ""], ["SENDER_NAME", null]]);

const RECIPIENT_COL  = "Recipient";
const EMAIL_SENT_COL = "Email Sent";

function processEmailsMonitored() {
  withHealthcheck_(processEmails)();
}

function myFunction() {
    console.log("Toto je výchozí funkce.");
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
    // console.log(`Try to fill template: Template = ${JSON.stringify(template, null, 2)}, Data = ${JSON.stringify(data, null, 2)}`);
    let template_string = JSON.stringify(template);

    // Token replacement
    template_string = template_string.replace(/{{[^{}]+}}/g, key => {
        return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
    });
    return JSON.parse(template_string);
}


/**
 * Escape cell data to make JSON safe
 * @see https://stackoverflow.com/a/9204218/1027723
 * @param {string} str to escape JSON special characters from
 * @return {string} escaped string
 */
function escapeData_(str) {
    // console.log(`Escaping data: ${str}, Type: ${typeof str}`);
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
        dataString += `*X-VS:${variableSymbol}`;
    }

    if (message) {
        dataString += `*MSG:${message}`;
    }

    return dataString;
}

/**
 * Consolidates data for QR code generation.
 * @param {object} qrCodeObj The QR code object with all necessary fields.
 * @param {Array} recipientRow The current recipient's row data from the realization sheet.
 * @param {Array} realizationHeaders The headers for the realization data.
 * @return {object} Consolidated QR code data with resolved variable symbol, amount, and message.
 */
function consolidateQrCodeData(qrCodeObj, recipientRow, realizationHeaders) {
    if (!qrCodeObj) {
        throw new Error("QR code object is required.");
    }

    // Validate that only one of each pair should be non-empty
    if (qrCodeObj.variableSymbol && qrCodeObj.variableSymbolColumn) {
        throw new Error("Only one of variableSymbol or variableSymbolColumn should be non-empty.");
    }

    if (qrCodeObj.amount && qrCodeObj.amountColumn) {
        throw new Error("Only one of amount or amountColumn should be non-empty.");
    }

    if (qrCodeObj.message && qrCodeObj.messageColumn) {
        throw new Error("Only one of message or messageColumn should be non-empty.");
    }

    // Resolve variable symbol
    let resolvedVariableSymbol = qrCodeObj.variableSymbol;
    if (!resolvedVariableSymbol && qrCodeObj.variableSymbolColumn) {
        const columnIdx = realizationHeaders.indexOf(qrCodeObj.variableSymbolColumn);
        if (columnIdx === -1) {
            throw new Error(`Column ${qrCodeObj.variableSymbolColumn} not found in realization data headers.`);
        }

        resolvedVariableSymbol = recipientRow[columnIdx];

        if (resolvedVariableSymbol == null) {
            throw new Error(`No value found in realization data for column ${qrCodeObj.variableSymbolColumn}.`);
        }
    }

    // Resolve amount
    let resolvedAmount = qrCodeObj.amount;
    if (!resolvedAmount && qrCodeObj.amountColumn) {
        const columnIdx = realizationHeaders.indexOf(qrCodeObj.amountColumn);
        if (columnIdx === -1) {
            throw new Error(`Column ${qrCodeObj.amountColumn} not found in realization data headers.`);
        }

        resolvedAmount = recipientRow[columnIdx];

        if (resolvedAmount == null) {
            throw new Error(`No value found in realization data for column ${qrCodeObj.amountColumn}.`);
        }
    }

    // Convert amount to number and validate
    if (resolvedAmount != null) {
        if (typeof resolvedAmount === 'string') {
            resolvedAmount = parseFloat(resolvedAmount);
            if (isNaN(resolvedAmount)) {
                const sourceInfo = qrCodeObj.amountColumn ? `column ${qrCodeObj.amountColumn}` : 'amount field';
                throw new Error(`Invalid amount value in ${sourceInfo} for QR code '${qrCodeObj.imageName}': ${qrCodeObj.amountColumn ? recipientRow[realizationHeaders.indexOf(qrCodeObj.amountColumn)] : qrCodeObj.amount}`);
            }
        } else if (typeof resolvedAmount !== 'number') {
            const sourceInfo = qrCodeObj.amountColumn ? `column ${qrCodeObj.amountColumn}` : 'amount field';
            throw new Error(`Amount must be a number in ${sourceInfo} for QR code '${qrCodeObj.imageName}': ${resolvedAmount}`);
        }
    } else {
        throw new Error(`Amount is required for QR code '${qrCodeObj.imageName}' - provide either 'amount' or 'amountColumn'.`);
    }

    // Resolve message
    let resolvedMessage = qrCodeObj.message;
    if (!resolvedMessage && qrCodeObj.messageColumn) {
        const columnIdx = realizationHeaders.indexOf(qrCodeObj.messageColumn);
        if (columnIdx === -1) {
            throw new Error(`Column ${qrCodeObj.messageColumn} not found in realization data headers.`);
        }

        resolvedMessage = recipientRow[columnIdx];

        if (resolvedMessage == null) {
            throw new Error(`No value found in realization data for column ${qrCodeObj.messageColumn}.`);
        }
    }

    return {
        ...qrCodeObj,
        variableSymbol: resolvedVariableSymbol,
        amount: resolvedAmount,
        message: resolvedMessage
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
        templateSource: templateSourceIdx !== -1 ? String(row[templateSourceIdx] || "") : ""
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
    console.log("Start insertQrCodesIntoEmail")
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
        console.error(`Failed to load realization document for topic "${plan.emailTopic}": ${error.message}`);
        return;
      }
    }

    const { headers: realizationHeaders, data: realizationData, qrCodes, sheet: realizationSheet } = realizationDoc;

    const recipientIdx = realizationHeaders.indexOf(RECIPIENT_COL);
    if (recipientIdx === -1) {
      console.error(`Recipient column (${RECIPIENT_COL}) not found in document ${plan.documentId}, sheet ${plan.sheetName}.`);
      return;
    }

    const conditionColIdx = realizationHeaders.indexOf(plan.conditionColumn);
    const sentColIdx = realizationHeaders.indexOf(plan.sentColumn);

    if (conditionColIdx === -1 || sentColIdx === -1) {
      console.error(`Columns "${plan.conditionColumn}" or "${plan.sentColumn}" not found in document ${plan.documentId}, sheet ${plan.sheetName}.`);
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
        console.error(`Failed to get template for ${recipient}: ${error.message}`);
        writeSentStatus(`${error.message} at ${new Date().toISOString()}`);
        return;
      }

      try {
        dataMapping = realizationHeaders.reduce((map, header, index) => {
          map[header] = recipientRow[index];
          return map;
        }, {});
      } catch (error) {
        console.error(`Failed to parse realization for ${recipient}: ${error.message}`);
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
        console.error(`Failed to generate QR code for ${recipient}: ${error.message}`);
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
        console.error(`Failed to fill template for ${recipient}: ${error.message}`);
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
        console.error(`Failed to inline images for ${recipient}: ${error.message}`);
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
        console.error(`Failed to send email to ${recipient}: ${error.message}`);
        writeSentStatus(`${error.message} at ${new Date().toISOString()}`);
      }
    });
  });
}

function processEmailsWithParams(planData, sheetsData, realizationSheetName) {
    // Get realization data from the specified sheet
    const realizationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(realizationSheetName);
    const realizationData = realizationSheet ? realizationSheet.getDataRange().getValues().slice(1) : []; // Skip headers
    const realizationHeaders = realizationSheet ? realizationSheet.getDataRange().getValues()[0] : []; // Get headers

    const recipientIdx = realizationHeaders.indexOf(MAIL_CONFIG.RECIPIENT_COL);
    if (recipientIdx === -1) {
        throw new Error(`Recipient column (${MAIL_CONFIG.RECIPIENT_COL}) not found in realization data.`);
    }

    realizationData.forEach((recipientRow, recipientIndex) => {
        const recipient = recipientRow[recipientIdx];

        planData.forEach(plan => {
            const conditionColIdx = realizationHeaders.indexOf(plan.conditionColumn);
            const sentColIdx = realizationHeaders.indexOf(plan.sentColumn);

            if (conditionColIdx === -1 || sentColIdx === -1) {
                throw new Error(`Columns ${plan.conditionColumn} or ${plan.sentColumn} not found in realization data.`);
            }

            const conditionValue = recipientRow[conditionColIdx];
            const sentValue = recipientRow[sentColIdx];

            console.log(
                `Parsed recipient row ${recipientIndex + 2}: ` +
                `Recipient = ${recipient}, Condition Value = ${conditionValue}, Sent Value = ${sentValue}`
            );

            if ((conditionValue === 1 && sentValue === "") || (conditionValue === 2)) {
                console.log(
                    `Preparing to send email for topic: ${plan.emailTopic} to recipient: ${recipient} at row ${recipientIndex + 2}.`
                );

                let sentStatus;
                let dataMapping;
                let consolidatedQrCode;
                let emailTemplate;
                let msgObj;
                let attachments;
                let inlineImages;

                console.log("Start looking for mail template.")
                try {
                    emailTemplate = getGmailTemplateFromDrafts_(plan.emailTopic);
                } catch (error) {
                    console.error(`Failed get template from drafts ${recipient}: ${error.message}`);
                    sentStatus = `${error.message} at ${Date()}`;
                    console.log("Stacktrace:", error.stack);
                    realizationSheet.getRange(recipientIndex + 2, sentColIdx + 1).setValue(sentStatus);
                    return null
                }

                console.log("Start preparing data mapping.")
                try {
                    // Create mapping of realization headers to their respective data
                    dataMapping = realizationHeaders.reduce((map, header, index) => {
                        map[header] = recipientRow[index];
                        return map;
                    }, {});
                } catch (error) {
                    console.error(`Failed to parse realisation ${recipient}: ${error.message}`);
                    sentStatus = `${error.message} at ${new Date().toISOString()}`;
                    console.log("Stacktrace:", error.stack);
                    realizationSheet.getRange(recipientIndex + 2, sentColIdx + 1).setValue(sentStatus);
                    return null
                }

                console.log("Start preparing qr code data mapping.")
                try {
                    // Include QR code mappings for specific email topic
                    if (sheetsData.qrCodes[plan.emailTopic]) {
                        sheetsData.qrCodes[plan.emailTopic].forEach(qrCodeObj => {
                            consolidatedQrCode = consolidateQrCodeData(qrCodeObj, recipientRow, realizationHeaders);
                            dataMapping[`${qrCodeObj.imageName}`] = `<img src="cid:${qrCodeObj.imageName}" alt="QR Code">`;
                        });
                    }

                    // Include global QR codes (available for all email topics)
                    if (sheetsData.qrCodes["GLOBAL"]) {
                        sheetsData.qrCodes["GLOBAL"].forEach(qrCodeObj => {
                            consolidatedQrCode = consolidateQrCodeData(qrCodeObj, recipientRow, realizationHeaders);
                            dataMapping[`${qrCodeObj.imageName}`] = `<img src="cid:${qrCodeObj.imageName}" alt="QR Code">`;
                        });
                    }
                } catch (error) {
                    console.error(`Failed to generate qr code ${recipient}: ${error.message}`);
                    sentStatus = `${error.message} at ${new Date().toISOString()}`;
                    console.log("Stacktrace:", error.stack);
                    realizationSheet.getRange(recipientIndex + 2, sentColIdx + 1).setValue(sentStatus);
                    return null
                }

                console.log("Start fill in template object.")
                try {
                    msgObj = fillInTemplateFromObject_(emailTemplate.message, {
                        ...dataMapping,
                        recipient,
                        conditionValue,
                        sentValue
                    });
                } catch (error) {
                    console.error(`Failed to fill template ${recipient}: ${error.message}`);
                    sentStatus = `${error.message} at ${new Date().toISOString()}`;
                    console.log("Stacktrace:", error.stack);
                    realizationSheet.getRange(recipientIndex + 2, sentColIdx + 1).setValue(sentStatus);
                    return null
                }

                console.log("Start inlining images.")
                try {
                    attachments = emailTemplate.attachments || [];
                    inlineImages = emailTemplate.inlineImages || {};

                    // Add QR codes and replace placeholders in email content for specific email topic
                    if (sheetsData.qrCodes[plan.emailTopic]) {
                        sheetsData.qrCodes[plan.emailTopic].forEach(qrCodeObj => {
                            // Only generate and insert QR codes that are actually needed in the email
                            if (emailTemplate.message.html.includes(`{{${qrCodeObj.imageName}}}`)) {
                                consolidatedQrCode = consolidateQrCodeData(qrCodeObj, recipientRow, realizationHeaders);
                                msgObj.html = insertQrCodesIntoEmail(msgObj.html, inlineImages, consolidatedQrCode);
                            }
                        });
                    }

                    // Add global QR codes and replace placeholders in email content
                    if (sheetsData.qrCodes["GLOBAL"]) {
                        sheetsData.qrCodes["GLOBAL"].forEach(qrCodeObj => {
                            // Only generate and insert global QR codes that are actually needed in the email
                            if (emailTemplate.message.html.includes(`{{${qrCodeObj.imageName}}}`)) {
                                consolidatedQrCode = consolidateQrCodeData(qrCodeObj, recipientRow, realizationHeaders);
                                msgObj.html = insertQrCodesIntoEmail(msgObj.html, inlineImages, consolidatedQrCode);
                            }
                        });
                    }
                } catch (error) {
                    console.error(`Failed to inline images ${recipient}: ${error.message}`);
                    sentStatus = `${error.message} at ${new Date().toISOString()}`;
                    console.log("Stacktrace:", error.stack);
                    realizationSheet.getRange(recipientIndex + 2, sentColIdx + 1).setValue(sentStatus);
                    return null
                }

                console.log("Start sending email.")
                try {
                    const emailOptions = {
                        from: plan.senderEmail,
                        cc:  plan.cc,
                        bcc: plan.bcc,
                        name: plan.senderName,
                        htmlBody: msgObj.html,
                        attachments: attachments,
                        inlineImages: inlineImages
                    };

                    GmailApp.sendEmail(recipient, msgObj.subject, msgObj.text, emailOptions);

                    sentStatus = new Date().toISOString();
                } catch (error) {
                    console.error(`Failed to send prepared email ${recipient}: ${error.message}`);
                    sentStatus = `${error.message} at ${new Date().toISOString()}`;
                    console.log("Stacktrace:", error.stack);
                    realizationSheet.getRange(recipientIndex + 2, sentColIdx + 1).setValue(sentStatus);
                    return null
                }

                // Update the "realizace" sheet with the status
                const range = realizationSheet.getRange(recipientIndex + 2, sentColIdx + 1);

                range.setValue(sentStatus);

                range.setNumberFormat("dd.MM.yyyy HH:mm:ss");
            }
        });
    });
}
