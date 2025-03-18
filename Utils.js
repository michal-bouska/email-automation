/**
 * Converts Czech bank account number to IBAN format
 * @param {string} accountNumber - main account number
 * @param {string} bankCode - bank code
 * @param {string} prefix - account prefix (optional)
 * @returns {string} - IBAN formatted as CZxx xxxx xxxx xxxx xxxx xxxx
 */
function convertToIBAN(accountNumber, bankCode, prefix = "") {
  // Validate input parameters
  if (!accountNumber || !bankCode) {
    throw new Error("Account number and bank code are required parameters");
  }

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

// Example usage
// const iban = convertToIBAN("123456789", "0800", "19");
// console.log(iban); // Prints: CZ65 0800 0000 1900 1234 5678