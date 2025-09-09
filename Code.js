const MAIL_CONFIG = loadConfig([["RECIPIENT_COL", "Recipient"], ["EMAIL_PLAN_SHEET", "plan"], ["EMAIL_LOG_SHEET", "realizace"], ["SENDER_EMAIL", null], ["CC_EMAIL", ""], ["SENDER_NAME", null]]);

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
    return JSON.parse(template_string);
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

function getGmailTemplateFromDrafts_(subject_line) {
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
        const allInlineImages = draft.getMessage().getAttachments({
            includeInlineImages: true,
            includeAttachments: false
        });
        const attachments = draft.getMessage().getAttachments({includeInlineImages: false});
        const htmlBody = msg.getBody();

        // Creates an inline image object with the image name as key
        // (can't rely on image index as array based on insert order)
        const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj), {});

        //Regexp searches for all img string positions with cid
        const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
        const matches = [...htmlBody.matchAll(imgexp)];

        //Initiates the allInlineImages object
        const inlineImagesObj = {};
        // built an inlineImagesObj from inline image matches
        matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

        return {
            message: {subject: subject_line, text: msg.getPlainBody(), html: htmlBody},
            attachments: attachments, inlineImages: inlineImagesObj
        };
    } catch (e) {
        Logger.error(`Error finding Gmail draft for subject '${subject_line}':`, e.message);
        Logger.error('Stacktrace:', e.stack);
        throw new Error(`Oops - can't find Gmail draft for subject '${subject_line}': ${e.message}`);
    }

    /**
     * Filter draft objects with the matching subject linemessage by matching the subject line.
     * @param {string} subject_line to search for draft message
     * @return {object} GmailDraft object
     */
    function subjectFilter_(subject_line) {
        return function (element) {
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
        Logger.error("QR code data object is missing.");
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
            Logger.error(`Failed to fetch QR code. HTTP response code ${response.getResponseCode()}.`);
            return null;
        }
    } catch (e) {
        Logger.error(`Error generating QR code: ${e.message}`);
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


const SHEETS_DATA = (() => {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const planSheet = spreadsheet.getSheetByName(MAIL_CONFIG.EMAIL_PLAN_SHEET);
    const qrCodeSheet = spreadsheet.getSheetByName("QRCodes");

    const planData = planSheet ? planSheet.getDataRange().getValues() : [];
    const qrCodeData = qrCodeSheet ? qrCodeSheet.getDataRange().getValues() : [];

    // Parse QR Code data into objects mapped by Email Topic and global QR codes
    const qrCodeHeaders = qrCodeData[0];
    const qrCodeObjects = qrCodeData.slice(1).reduce((map, row) => {
        const qrCodeObj = {
            emailTopic: row[qrCodeHeaders.indexOf("EmailTopic")],
            imageName: row[qrCodeHeaders.indexOf("ImageName")],
            accountNumber: row[qrCodeHeaders.indexOf("AccountNumber")],
            bankCode: row[qrCodeHeaders.indexOf("BankCode")],
            currency: row[qrCodeHeaders.indexOf("Currency")],
            amount: row[qrCodeHeaders.indexOf("Amount")],
            amountColumn: row[qrCodeHeaders.indexOf("AmountColumn")],
            variableSymbol: row[qrCodeHeaders.indexOf("VariableSymbol")],
            variableSymbolColumn: row[qrCodeHeaders.indexOf("VariableSymbolColumn")],
            message: row[qrCodeHeaders.indexOf("Message")],
            messageColumn: row[qrCodeHeaders.indexOf("MessageColumn")],
            size: row[qrCodeHeaders.indexOf("Size")]
        };

        // Skip rows without imageName
        if (!qrCodeObj.imageName) {
            return map;
        }

        if (qrCodeObj.emailTopic) {
            // QR code for specific email topic
            if (!map[qrCodeObj.emailTopic]) {
                map[qrCodeObj.emailTopic] = [];
            }
            map[qrCodeObj.emailTopic].push(qrCodeObj);
        } else {
            // Global QR code (empty EmailTopic) - add to special "GLOBAL" key
            if (!map["GLOBAL"]) {
                map["GLOBAL"] = [];
            }
            map["GLOBAL"].push(qrCodeObj);
        }

        return map;
    }, {});

    Logger.log("Loaded QR Code Data: %s", JSON.stringify(qrCodeObjects, null, 2));

    return {
        plan: planData,
        qrCodes: qrCodeObjects
    };
})();

/**
 * Transforms plan data into objects with validation and groups them by realization sheet.
 * Assumes the first row contains headers "Email Topic", "Column Condition To Send", "Column Sent",
 * and optionally "Sender Email", "CC", "BCC", "Sender Name", "Realization Sheet".
 * Checks that "Column Condition To Send" and "Column Sent" contain valid column names ([a-z]+).
 * @returns {Object} Object with realization sheet names as keys and arrays of plan objects as values.
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

    // Optional email configuration columns
    const senderEmailIdx = headers.indexOf("Sender Email");
    const ccIdx = headers.indexOf("CC");
    const bccIdx = headers.indexOf("BCC");
    const senderNameIdx = headers.indexOf("Sender Name");
    const realizationSheetIdx = headers.indexOf("Realization Sheet");

    if (emailTopicIdx === -1 || conditionColumnIdx === -1 || sentDateColumnIdx === -1) {
        throw new Error("Required headers are missing in the plan data.");
    }

    const plans = planData.slice(1).map(row => {
        return {
            emailTopic: row[emailTopicIdx],
            conditionColumn: row[conditionColumnIdx],
            sentColumn: row[sentDateColumnIdx],
            // Email configuration with fallback to MAIL_CONFIG defaults
            senderEmail: senderEmailIdx !== -1 && row[senderEmailIdx] ? row[senderEmailIdx] : MAIL_CONFIG.SENDER_EMAIL,
            cc: ccIdx !== -1 && row[ccIdx] ? row[ccIdx] : MAIL_CONFIG.CC_EMAIL,
            bcc: bccIdx !== -1 && row[bccIdx] ? row[bccIdx] : null,
            senderName: senderNameIdx !== -1 && row[senderNameIdx] ? row[senderNameIdx] : MAIL_CONFIG.SENDER_NAME,
            // Realization sheet with fallback to MAIL_CONFIG default
            realizationSheet: realizationSheetIdx !== -1 && row[realizationSheetIdx] ? row[realizationSheetIdx] : MAIL_CONFIG.EMAIL_LOG_SHEET
        };
    });

    // Group plans by realization sheet
    const groupedPlans = plans.reduce((groups, plan) => {
        const sheetName = plan.realizationSheet;
        if (!groups[sheetName]) {
            groups[sheetName] = [];
        }
        groups[sheetName].push(plan);
        return groups;
    }, {});

    return groupedPlans;
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
    const groupedPlans = parsePlanData();

    // Process each realization sheet separately
    Object.keys(groupedPlans).forEach(realizationSheetName => {
        const plansForSheet = groupedPlans[realizationSheetName];
        Logger.log(`Processing ${plansForSheet.length} plans for realization sheet: ${realizationSheetName}`);

        processEmailsWithParams(plansForSheet, SHEETS_DATA, realizationSheetName);
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

            Logger.log(
                `Parsed recipient row ${recipientIndex + 2}: ` +
                `Recipient = ${recipient}, Condition Value = ${conditionValue}, Sent Value = ${sentValue}`
            );

            if ((conditionValue === 1 && sentValue === "") || (conditionValue === 2)) {
                Logger.log(
                    `Preparing to send email for topic: ${plan.emailTopic} to recipient: ${recipient} at row ${recipientIndex + 2}.`
                );

                let sentStatus;
                let dataMapping;
                let consolidatedQrCode;
                let emailTemplate;
                let msgObj;
                let attachments;
                let inlineImages;

                Logger.log("Start looking for mail template.")
                try {
                    emailTemplate = getGmailTemplateFromDrafts_(plan.emailTopic);
                } catch (error) {
                    Logger.error(`Failed get template from drafts ${recipient}: ${error.message}`);
                    sentStatus = `${error.message} at ${Date()}`;
                    Logger.log("Stacktrace:", error.stack);
                    realizationSheet.getRange(recipientIndex + 2, sentColIdx + 1).setValue(sentStatus);
                    return null
                }

                Logger.log("Start preparing data mapping.")
                try {
                    // Create mapping of realization headers to their respective data
                    dataMapping = realizationHeaders.reduce((map, header, index) => {
                        map[header] = recipientRow[index];
                        return map;
                    }, {});
                } catch (error) {
                    Logger.error(`Failed to parse realisation ${recipient}: ${error.message}`);
                    sentStatus = `${error.message} at ${new Date().toISOString()}`;
                    Logger.log("Stacktrace:", error.stack);
                    realizationSheet.getRange(recipientIndex + 2, sentColIdx + 1).setValue(sentStatus);
                    return null
                }

                Logger.log("Start preparing qr code data mapping.")
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
                    Logger.error(`Failed to generate qr code ${recipient}: ${error.message}`);
                    sentStatus = `${error.message} at ${new Date().toISOString()}`;
                    Logger.log("Stacktrace:", error.stack);
                    realizationSheet.getRange(recipientIndex + 2, sentColIdx + 1).setValue(sentStatus);
                    return null
                }

                Logger.log("Start fill in template object.")
                try {
                    msgObj = fillInTemplateFromObject_(emailTemplate.message, {
                        ...dataMapping,
                        recipient,
                        conditionValue,
                        sentValue
                    });
                } catch (error) {
                    Logger.error(`Failed to fill template ${recipient}: ${error.message}`);
                    sentStatus = `${error.message} at ${new Date().toISOString()}`;
                    Logger.log("Stacktrace:", error.stack);
                    realizationSheet.getRange(recipientIndex + 2, sentColIdx + 1).setValue(sentStatus);
                    return null
                }

                Logger.log("Start inlining images.")
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
                    Logger.error(`Failed to inline images ${recipient}: ${error.message}`);
                    sentStatus = `${error.message} at ${new Date().toISOString()}`;
                    Logger.log("Stacktrace:", error.stack);
                    realizationSheet.getRange(recipientIndex + 2, sentColIdx + 1).setValue(sentStatus);
                    return null
                }

                Logger.log("Start sending email.")
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
                    Logger.error(`Failed to send prepared email ${recipient}: ${error.message}`);
                    sentStatus = `${error.message} at ${new Date().toISOString()}`;
                    Logger.log("Stacktrace:", error.stack);
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
