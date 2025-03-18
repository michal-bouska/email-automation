/**
 * Fetch transactions from Fio Bank API using Google Apps Script
 * This script demonstrates how to connect to Fio Bank API and retrieve account transactions
 * with additional parsing for specific transaction keys in recipient notes
 */

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
 * Get transactions for a specific date range
 * @param {string} dateFrom - Start date in format YYYY-MM-DD
 * @param {string} dateTo - End date in format YYYY-MM-DD
 * @return {Object} JSON response with transactions
 */
function getTransactions(dateFrom, dateTo) {
  // Format the API URL with date parameters
  const url = `https://www.fio.cz/ib_api/rest/periods/${FIO_API_TOKEN}/${dateFrom}/${dateTo}/transactions.json`;

  try {
    // Make the HTTP request
    const response = UrlFetchApp.fetch(url, {
      'method': 'get',
      'muteHttpExceptions': true
    });

    // Parse and return JSON response
    return JSON.parse(response.getContentText());
  } catch (error) {
    Logger.log('Error fetching data from Fio API: ' + error);
    return null;
  }
}

/**
 * Get the most recent transactions (last 7 days by default)
 * @return {Object} JSON response with recent transactions
 */
function getRecentTransactions() {
  // Calculate dates for the last 7 days
  const today = new Date();
  const lastWeek = new Date();
  lastWeek.setDate(today.getDate() - 7);

  // Format dates as YYYY-MM-DD
  const dateFrom = Utilities.formatDate(lastWeek, 'Europe/Prague', 'yyyy-MM-dd');
  const dateTo = Utilities.formatDate(today, 'Europe/Prague', 'yyyy-MM-dd');

  return getTransactions(dateFrom, dateTo);
}

/**
 * Parse transaction recipient notes to extract specific keys
 * @param {string} note - The recipient note text to parse
 * @return {Object} Object containing parsed keys and their count
 */
function parseTransactionNote(note) {
  if (!note) {
    return { keys: '', count: 0 };
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
 * Process transactions and write only new ones to a Google Sheet
 * Include parsing of transaction notes
 * @param {string} sheetId - ID of the Google Sheet
 * @param {string} sheetName - Name of the sheet tab
 */
function writeTransactionsToSheet(sheetId, sheetName) {
  // Get transactions data
  const transactionsData = getRecentTransactions();

  if (!transactionsData || !transactionsData.accountStatement || !transactionsData.accountStatement.transactionList) {
    Logger.log('No transaction data available');
    return;
  }

  // Get the transactions array
  const transactions = transactionsData.accountStatement.transactionList.transaction;

  // Open the specified Google Sheet
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Sheet "${sheetName}" not found in the spreadsheet`);
    return;
  }

  // Prepare headers with new columns for parsed keys and count
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

  // Get existing transaction IDs to avoid duplicates
  // Skip header row if it exists
  const startRow = existingData.length > 0 ? 1 : 0;
  const existingTransactionIds = new Set();

  // Find the transaction ID column
  const idColumnIndex = headers.indexOf('Transaction ID');

  // Collect existing transaction IDs
  for (let i = startRow; i < existingData.length; i++) {
    if (existingData[i].length > idColumnIndex) {
      const transactionId = existingData[i][idColumnIndex];
      if (transactionId) {
        existingTransactionIds.add(transactionId);
      }
    }
  }

  // Counter for new transactions
  let newTransactionsCount = 0;

  // Process and append only new transactions with the specific variable symbol
  transactions.forEach(transaction => {
    const transactionId = transaction.id?.value || '';
    const variableSymbol = transaction.variableSymbol?.value || '';

    // Skip if this transaction ID already exists in the sheet
    if (transactionId && existingTransactionIds.has(transactionId)) {
      return;
    }

    // Only process transactions with the specific variable symbol
    if (variableSymbol === VARIABLE_SYMBOL_KEY) {
      const message = transaction.message?.value || '';

      // Parse the transaction note
      const parsedNote = parseTransactionNote(message);

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
      newTransactionsCount++;
    }
  });

  // Format the sheet (only if we have data)
  if (existingData.length > 0 || newTransactionsCount > 0) {
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.autoResizeColumns(1, headers.length);
  }

  Logger.log(`Added ${newTransactionsCount} new transactions with variable symbol ${VARIABLE_SYMBOL_KEY}`);
}

/**
 * Create a trigger to automatically fetch transactions every 15 minutes
 */
function create15MinTrigger() {
  // Delete any existing triggers with the same function name
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runUpdate') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Create a new trigger to run every 15 minutes
  ScriptApp.newTrigger('runUpdate')
      .timeBased()
      .everyMinutes(15)
      .create();
}

/**
 * Function that will be called by the 15-minute trigger
 * Update this with your specific sheet ID and name
 */
function runUpdate() {
  const SHEET_ID = 'your_google_sheet_id_here';
  const SHEET_NAME = 'Fio Transactions';

  writeTransactionsToSheet(SHEET_ID, SHEET_NAME);
}

/**
 * Function to manually run the script from the script editor
 */
function manualRun() {
  runUpdate();
}