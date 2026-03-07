/**
 * Converts Czech bank account number to IBAN format
 * @param {string|number} accountNumber - main account number
 * @param {string|number} bankCode - bank code
 * @param {string|number} prefix - account prefix (optional)
 * @returns {string} - IBAN formatted as CZxx xxxx xxxx xxxx xxxx xxxx
 */
function convertToIBAN(accountNumber, bankCode, prefix = "") {
  // Validate input parameters
  if (accountNumber === undefined || accountNumber === null ||
      bankCode === undefined || bankCode === null) {
    throw new Error("Account number and bank code are required parameters");
  }

  // Convert all parameters to strings
  accountNumber = String(accountNumber);
  bankCode = String(bankCode);
  prefix = prefix !== undefined && prefix !== null ? String(prefix) : "";

  // Remove spaces and other non-numeric characters
  accountNumber = accountNumber.replace(/\D/g, "");
  prefix = prefix.replace(/\D/g, "");
  bankCode = bankCode.replace(/\D/g, "");

  // Pad with zeros from left to correct length
  accountNumber = accountNumber.padStart(10, "0");
  prefix = prefix.padStart(6, "0");
  bankCode = bankCode.padStart(4, "0");

  // BBAN (Basic Bank Account Number) format for Czech Republic
  const bban = bankCode + prefix + accountNumber;

  // Convert country code "CZ" to numeric format (C=3, Z=35) -> "32635"
  const countryCode = "CZ";
  const countryCodeNum = "3235";

  // Add "00" at the end (check digits, initially set to 00)
  const numericRepresentation = bban + countryCodeNum + "00";

  // Calculate modulo 97 according to ISO 7064
  let checksum = 98 - (modulo97(numericRepresentation) % 97);
  checksum = checksum.toString().padStart(2, "0");

  // Assemble the final IBAN
  const iban = countryCode + checksum + bban;

  // Format IBAN with spaces for better readability
  return formatIBAN(iban);
}

/**
 * Calculate modulo 97 for large numbers (ISO 7064 standard)
 * @param {string} numStr - input string of numbers
 * @returns {number} - modulo 97 result
 */
function modulo97(numStr) {
  // For large numbers that could cause overflow, we use iterative calculation
  let remainder = 0;

  for (let i = 0; i < numStr.length; i++) {
    remainder = (remainder * 10 + parseInt(numStr[i])) % 97;
  }

  return remainder;
}

/**
 * Formats IBAN into readable format with spaces every 4 characters
 * @param {string} iban - IBAN without spaces
 * @returns {string} - IBAN with spaces
 */
function formatIBAN(iban) {
  return iban.match(/.{1,4}/g).join(" ");
}

// Example usage with different parameter types
// const iban1 = convertToIBAN("123456789", "0800", "19");      // All strings
// const iban2 = convertToIBAN(123456789, 800, 19);             // All integers
// const iban3 = convertToIBAN("123456789", 800, 19);           // Mixed types
// console.log(iban1); // Prints: CZ65 0800 0000 1900 1234 5678

/**
 * Sends a ping to Healthchecks.io. Failures are logged but never block execution.
 * @param {string} suffix URL suffix: "" for success, "/start" for start, "/fail" for failure.
 * @param {string} body Optional POST body (error details on /fail).
 * @param {string} propertyName Script Property name for the ping URL (default: HEALTHCHECKS_PING_URL).
 */
function pingHealthcheck_(suffix, body, propertyName) {
  propertyName = propertyName || 'HEALTHCHECKS_PING_URL';
  const pingUrl = PropertiesService.getScriptProperties().getProperty(propertyName);
  if (!pingUrl) {
    Logger.log(`${propertyName} not configured, skipping healthcheck ping.`);
    return;
  }

  const url = pingUrl.replace(/\/+$/, '') + suffix;

  try {
    const options = {
      method: body ? 'post' : 'get',
      muteHttpExceptions: true
    };

    if (body) {
      options.payload = body;
      options.contentType = 'text/plain';
    }

    const response = UrlFetchApp.fetch(url, options);
    Logger.log(`Healthcheck ping ${suffix || '/success'}: HTTP ${response.getResponseCode()}`);
  } catch (error) {
    Logger.log(`Healthcheck ping failed (non-blocking): ${error.message}`);
  }
}

/**
 * Wraps a function with Healthchecks.io start/success/fail pings.
 * @param {Function} fn The function to wrap.
 * @param {string} propertyName Script Property name for the ping URL.
 * @return {Function} Wrapped function.
 */
function withHealthcheck_(fn, propertyName) {
  return function(...args) {
    pingHealthcheck_('/start', null, propertyName);
    try {
      const result = fn.apply(this, args);
      pingHealthcheck_('', null, propertyName);
      return result;
    } catch (error) {
      pingHealthcheck_('/fail', `${error.message}\n\n${error.stack || ''}`, propertyName);
      throw error;
    }
  };
}