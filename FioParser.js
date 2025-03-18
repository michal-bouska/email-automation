/**
 * Fetch transactions from Fio Bank API using Google Apps Script
 * This script connects to Fio Bank API and retrieves account transactions since the last fetch
 * with additional parsing for specific transaction keys in recipient notes
 */

/**
 * Function to manually run the script from the script editor
 */
function manualRun() {
    runUpdate();
}

// Function to get Fio token from secure storage
function getFioToken() {
    // Try to get token from Properties Service
    const scriptProperties = PropertiesService.getScriptProperties();
    const token = scriptProperties.getProperty('FIO_API_TOKEN');

    if (!token) {
        Logger.log('Fio API token not found in Properties Service');
        return 'your_fio_token_here'; // Default fallback value
    }

    return token;
}

// Variable symbol key to filter transactions
const VARIABLE_SYMBOL_KEY = '72405';

/**
 * Get transactions since the last fetch from Fio Bank API
 * The bank automatically tracks the last fetch on their server
 * @return {Object} JSON response with transactions
 */
function getLastTransactions() {
    // Format the API URL for last transactions endpoint
    const url = `https://www.fio.cz/ib_api/rest/last/${getFioToken()}/transactions.json`;
    Logger.log("URL: " + url)

    let response_value;
    try {
        // Make the HTTP request
        const response = UrlFetchApp.fetch(url, {
            'method': 'get',
            'muteHttpExceptions': true
        });

        response_value = response.getContentText()

        // Parse and return JSON response
        return JSON.parse(response_value);
    } catch (error) {
        Logger.log("Resp: " + response_value)
        Logger.log('Error fetching data from Fio API: ' + error);
        return null;
    }
}

/**
 * Parse transaction recipient notes to extract specific keys
 * @param {string} note - The recipient note text to parse
 * @return {Object} Object containing parsed keys and their count
 */
function parseTransactionNote(note) {
    if (!note) {
        return {keys: '', count: 0};
    }

    // Replace all non-numeric characters with spaces
    const cleanedNote = note.replace(/[^0-9]/g, ' ');

    // Split by spaces and filter out empty strings
    const parts = cleanedNote.split(' ').filter(part => part.trim() !== '');

    // Keep only strings with exactly 8 characters (keys)
    const validKeys = parts.filter(part => part.length === 8);

    return {
        keys: validKeys.join(','),
        count: validKeys.length
    };
}

/**
 * Process a single transaction and write it to the sheet
 * @param {Object} transaction - The transaction object from Fio API
 * @param {Object} sheet - The Google Sheet to write to
 * @param {string} variableSymbolKey - The variable symbol to filter by
 * @return {boolean} - Whether the transaction was processed and added to sheet
 */
function processTransaction(transaction, sheet, variableSymbolKey) {
  const variableSymbol = transaction.variableSymbol?.value || '';
  const message = transaction.message?.value || '';
  const transactionId = transaction.id?.value || '';

  let parsedNote;

  // Parse the transaction note
  if (variableSymbol === variableSymbolKey) {
    parsedNote = parseTransactionNote(message);
  } else {
    parsedNote = {keys: '', count: 0};
  }

  const row = [
    transaction.date?.value || '',
    transaction.amount?.value || '',
    transaction.currency?.value || '',
    transaction.accountNumber?.value || '',
    transaction.bankCode?.value || '',
    transaction.senderName?.value || '',
    transaction.type?.value || '',
    message,
    variableSymbol,
    parsedNote.keys,
    parsedNote.count,
    transactionId
  ];

  sheet.appendRow(row);

  // Return true if it was a variable symbol match, false otherwise
  return variableSymbol === variableSymbolKey;
}

/**
 * Process new transactions and write them to a Google Sheet
 * Include parsing of transaction notes
 * @param {string} sheetName - Name of the sheet tab
 */
function writeTransactionsToSheet(sheetName) {
  // Get new transactions data
  const transactionsData = getLastTransactions();

  if (!transactionsData || !transactionsData.accountStatement || !transactionsData.accountStatement.transactionList) {
    Logger.log('No new transaction data available');
    return;
  }

  // Get the transactions array
  const transactions = transactionsData.accountStatement.transactionList.transaction || [];

  if (transactions.length === 0) {
    Logger.log('No new transactions found');
    return;
  }

  // Open the specified Google Sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Sheet "${sheetName}" not found in the spreadsheet`);
    return;
  }

  // Prepare headers with columns for parsed keys and count
  const headers = [
    'Date', 'Amount', 'Currency', 'Account', 'Bank Code',
    'Sender Name', 'Transaction Type', 'Message', 'Variable Symbol', 'Parsed Keys', 'Keys Count', 'Transaction ID'
  ];

  // Check if the sheet is empty and add headers if needed
  const existingData = sheet.getDataRange().getValues();
  if (existingData.length === 0) {
    sheet.appendRow(headers);
  } else if (existingData[0].length < headers.length) {
    // Headers exist but may be missing our new columns - update them
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  // Counter for new transactions
  let newTransactionsCount = 0;
  let filteredTransactionsCount = 0;

  // Process each transaction
  transactions.forEach(transaction => {
    newTransactionsCount++;

    // Use the named function to process the transaction
    const wasProcessed = processTransaction(transaction, sheet, VARIABLE_SYMBOL_KEY);

    if (wasProcessed) {
      filteredTransactionsCount++;
    }
  });

  // Format the sheet (only if we have data)
  if (existingData.length === 0 || filteredTransactionsCount > 0) {
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.autoResizeColumns(1, headers.length);
  }

  Logger.log(`Total new transactions: ${newTransactionsCount}`);
  Logger.log(`Added ${filteredTransactionsCount} new transactions with variable symbol ${VARIABLE_SYMBOL_KEY}`);
}


/**
 * Create a trigger to automatically fetch transactions every day
 */
function createDailyTrigger() {
    // Delete any existing triggers with the same function name
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === 'runUpdate') {
            ScriptApp.deleteTrigger(triggers[i]);
        }
    }

    // Create a new trigger to run daily
    ScriptApp.newTrigger('runUpdate')
        .timeBased()
        .everyDays(1)
        .atHour(6) // Run at 6 AM
        .create();
}

/**
 * Function that will be called by the trigger
 */
function runUpdate() {
    const SHEET_NAME = 'FioSync';
    writeTransactionsToSheet(SHEET_NAME);
}

