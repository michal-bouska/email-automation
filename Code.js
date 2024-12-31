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
 * @OnlyCurrentDoc
*/
 
/**
 * Change these to match the column names you are using for email 
 * recipient addresses and email sent column.
*/
const RECIPIENT_COL  = "Recipient";
const EMAIL_SENT_COL = "Email Sent";

function myFunction() {
  Logger.log("Toto je výchozí funkce.");
  processEmails();
}
 
/** 
 * Creates the menu item "Mail Merge" for user to run scripts on drop-down.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
      .addItem('Send Emails', 'sendEmails')
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

function getGmailTemplateFromDrafts_(subject_line){
  try {
    // get drafts
    const drafts = GmailApp.getDrafts();
    // filter the drafts that match subject line
    const draft = drafts.filter(subjectFilter_(subject_line))[0];
    // get the message object
    const msg = draft.getMessage();

    // Handles inline images and attachments so they can be included in the merge
    // Based on https://stackoverflow.com/a/65813881/1027723
    // Gets all attachments and inline image attachments
    const allInlineImages = draft.getMessage().getAttachments({includeInlineImages: true,includeAttachments:false});
    const attachments = draft.getMessage().getAttachments({includeInlineImages: false});
    const htmlBody = msg.getBody(); 

    // Creates an inline image object with the image name as key 
    // (can't rely on image index as array based on insert order)
    const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj) ,{});

    //Regexp searches for all img string positions with cid
    const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
    const matches = [...htmlBody.matchAll(imgexp)];

    //Initiates the allInlineImages object
    const inlineImagesObj = {};
    // built an inlineImagesObj from inline image matches
    matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

    return {message: {subject: subject_line, text: msg.getPlainBody(), html:htmlBody}, 
            attachments: attachments, inlineImages: inlineImagesObj };
  } catch(e) {
    throw new Error("Oops - can't find Gmail draft");
  }

  /**
   * Filter draft objects with the matching subject linemessage by matching the subject line.
   * @param {string} subject_line to search for draft message
   * @return {object} GmailDraft object
  */
  function subjectFilter_(subject_line){
    return function(element) {
      if (element.getMessage().getSubject() === subject_line) {
        return element;
      }
    }
  }
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
  let dataString = `SPD*1.0*ACC:${accountNumber}/${bankCode}*AM:${amount.toFixed(2)}*CC:${currency}`;

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


const SHEETS_DATA = (() => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const planSheet = spreadsheet.getSheetByName("plan");
  const realizationSheet = spreadsheet.getSheetByName("realizace");
  const qrCodeSheet = spreadsheet.getSheetByName("QRCodes");

  const planData = planSheet ? planSheet.getDataRange().getValues() : [];
  const realizationData = realizationSheet ? realizationSheet.getDataRange().getValues() : [];
  const qrCodeData = qrCodeSheet ? qrCodeSheet.getDataRange().getValues() : [];

  // Parse QR Code data into objects mapped by Email Topic
  const qrCodeHeaders = qrCodeData[0];
  const qrCodeObjects = qrCodeData.slice(1).reduce((map, row) => {
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

  Logger.log("Loaded QR Code Data: %s", JSON.stringify(qrCodeObjects, null, 2));

  return {
    plan: planData,
    realization: realizationData,
    qrCodes: qrCodeObjects
  };
})();

/**
 * Transforms plan data into objects with validation.
 * Assumes the first row contains headers "Email Topic", "Column Condition To Send", and "Column Sent".
 * Checks that "Column Condition To Send" and "Column Sent" contain valid column names ([a-z]+).
 * @returns {Array<Object>} Array of parsed objects.
 */
function parsePlanData() {
  const planData = SHEETS_DATA.plan;
  if (!planData || planData.length < 2) {
    throw new Error("Plan data is missing or insufficient rows.");
  }

  const headers = planData[0];
  const emailTopicIdx = headers.indexOf("Email Topic");
  const conditionColumnIdx = headers.indexOf("Column Condition To Send");
  const sentDateColumnIdx = headers.indexOf("Column Sent");

  if (emailTopicIdx === -1 || conditionColumnIdx === -1 || sentDateColumnIdx === -1) {
    throw new Error("Required headers are missing in the plan data.");
  }

  const isValidColumnName = name => /^[A-Z]+$/.test(name);

  return planData.slice(1).map(row => {
    return {
      emailTopic: row[emailTopicIdx],
      conditionColumn: row[conditionColumnIdx],
      sentColumn: row[sentDateColumnIdx]
    };
  });
}

/**
 * Processes emails to send based on plan and realization data.
 */
function processEmails() {
  const planData = parsePlanData();
  const realizationData = SHEETS_DATA.realization.slice(1); // Skip headers
  const realizationHeaders = SHEETS_DATA.realization[0]; // Get headers
  const realizationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("realizace");

  const recipientIdx = realizationHeaders.indexOf(RECIPIENT_COL);
  if (recipientIdx === -1) {
    throw new Error(`Recipient column (${RECIPIENT_COL}) not found in realization data.`);
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

      if (conditionValue === 1 && sentValue === "") {
        console.log(
          `Preparing to send email for topic: ${plan.emailTopic} to recipient: ${recipient} at row ${recipientIndex + 2}.`
        );

        let sentStatus;

        try {
          const emailTemplate = getGmailTemplateFromDrafts_(plan.emailTopic);
          const msgObj = fillInTemplateFromObject_(emailTemplate.message, {
            recipient,
            conditionValue,
            sentValue
          });

          let attachments = emailTemplate.attachments || [];
          let inlineImages = emailTemplate.inlineImages || {};

          // Add QR codes to inlineImages
          if (SHEETS_DATA.qrCodes[plan.emailTopic]) {
            SHEETS_DATA.qrCodes[plan.emailTopic].forEach(qrCodeObj => {
              const consolidatedQrCode = consolidateQrCodeData(qrCodeObj, realizationData, realizationHeaders);
              const qrCodeBlob = generateQrCodeBlob(consolidatedQrCode);
              if (qrCodeBlob) {
                inlineImages[consolidatedQrCode.imageName] = qrCodeBlob;
              }
            });
          }

          GmailApp.sendEmail(recipient, msgObj.subject, msgObj.text, {
            htmlBody: msgObj.html,
            attachments: attachments,
            inlineImages: inlineImages
          });

          sentStatus = new Date(); // Store current date and time
        } catch (error) {
          console.error(`Failed to send email to ${recipient}: ${error.message}`);
          sentStatus = error.message; // Store error message
        }

        // Update the "realizace" sheet with the status
        realizationSheet.getRange(recipientIndex + 2, sentColIdx + 1).setValue(sentStatus);
      }
    });
  });
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
