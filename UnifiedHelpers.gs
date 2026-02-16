/**
* ============================================================================
* UNIFIED HELPERS - SHARED UTILITY FUNCTIONS
* ============================================================================
* 
* Consolidates all duplicate helper functions into a single module.
* Used by all reports and sync functions.
* 
* FUNCTIONS INCLUDED:
* - Column/Row Utilities (columnToLetter, letterToColumn)
* - Date Utilities (formatDateToYYYYMMM, parseDate, getFiscalYear, getFYString)
* - Numeric Utilities (parseNumeric, formatCurrency)
* - String Utilities (cleanString, isValidEmail)
* - Array Utilities (groupBy, sortBy)
* - Validation Utilities (isValidDate, isNonEmptyString)
* 
* USAGE:
* Import this file and use any helper function directly.
* No initialization required.
* 
* ============================================================================
*/

/**
 * ============================================================================
 * COLUMN & ROW UTILITIES
 * ============================================================================
 */

/**
 * Convert column number to Excel letter (1 -> A, 27 -> AA)
 * @param {number} column - Column number (1-based)
 * @returns {string} Column letter
 */
function columnToLetter(column) {
  let temp, letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (temp - column) / 26 + column;
    column = Math.floor(column);
  }
  return letter;
}

/**
 * Convert Excel letter to column number (A -> 1, AA -> 27)
 * @param {string} letter - Column letter
 * @returns {number} Column number
 */
function letterToColumn(letter) {
  let column = 0;
  const length = letter.length;
  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

/**
 * ============================================================================
 * DATE UTILITIES
 * ============================================================================
 */

/**
 * Format date to YYYY-MMM format (e.g., "2024-Jan")
 * @param {Date|string} date - Date to format
 * @returns {string} Formatted date string
 */
function formatDateToYYYYMMM(date) {
  if (!(date instanceof Date)) {
    if (typeof date === 'string') {
      date = new Date(date);
    } else {
      return String(date);
    }
  }
  
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return String(date);
  }
  
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  return `${date.getFullYear()}-${months[date.getMonth()]}`;
}

/**
 * Format date to YYYY-MM format (e.g., "2024-01")
 * @param {Date} date - Date to format
 * @returns {string} Formatted date string
 */
function formatDateToYYYYMM(date) {
  if (!(date instanceof Date)) return String(date);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  return `${year}-${month}`;
}

/**
 * Format date to ISO string (YYYY-MM-DD)
 * @param {Date} date - Date to format
 * @returns {string} ISO date string
 */
function formatDateToISO(date) {
  if (!(date instanceof Date)) return String(date);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * Get Fiscal Year string from date
 * @param {Date} date - Date to check
 * @param {number} fyStartMonth - Start month of FY (0-11, default 6 = July)
 * @returns {string} Fiscal year string (e.g., "FY 24/25 Q1")
 */
function getFiscalYear(date, fyStartMonth = 6) {
  if (!(date instanceof Date)) {
    date = new Date(date);
  }
  
  const year = date.getFullYear();
  const month = date.getMonth();
  
  const fyStart = (month >= fyStartMonth) ? year : year - 1;
  const shortYearStart = fyStart.toString().slice(-2);
  const shortYearEnd = (fyStart + 1).toString().slice(-2);
  
  let relativeMonth = month - fyStartMonth;
  if (relativeMonth < 0) relativeMonth += 12;
  const quarter = Math.floor(relativeMonth / 3) + 1;
  
  return `FY ${shortYearStart}/${shortYearEnd} Q${quarter}`;
}

/**
 * Get simple Fiscal Year (e.g., "FY 24/25")
 * @param {Date} date - Date to check
 * @param {number} fyStartMonth - Start month of FY (0-11)
 * @returns {string} Simple FY string
 */
function getFYString(date, fyStartMonth = 6) {
  if (!(date instanceof Date)) date = new Date(date);
  const year = date.getFullYear();
  const fyStart = (date.getMonth() >= fyStartMonth) ? year : year - 1;
  return `FY ${fyStart.toString().slice(-2)}/${(fyStart + 1).toString().slice(-2)}`;
}

/**
 * Find column index by date in headers
 * @param {Array} headers - Array of header strings
 * @param {Date} targetDate - Target date to find
 * @returns {number} Column index or -1
 */
function findDateColumnIndex(headers, targetDate) {
  const tStr = formatDateToYYYYMMM(targetDate);
  return headers.findIndex(h => formatDateToYYYYMMM(h) === tStr);
}

/**
 * Parse flexible date string
 * @param {string} dateString - Date string to parse
 * @returns {Date} Parsed date or epoch if invalid
 */
function parseFlexibleDate(dateString) {
  if (!dateString) return new Date(0);
  let normalized = dateString.toString().trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(normalized)) normalized += 'T00:00:00.000Z';
  const parsed = Date.parse(normalized);
  return isNaN(parsed) ? new Date(0) : new Date(parsed);
}

/**
 * Check if date is valid
 * @param {*} dateVal - Value to check
 * @returns {boolean} True if valid date
 */
function isValidDate(dateVal) {
  if (dateVal === null || dateVal === undefined || dateVal === "") return false;
  const strVal = dateVal.toString().trim().toLowerCase();
  if (["n/a", "unknown", "blank", ""].includes(strVal)) return false;
  
  const d = new Date(dateVal);
  return d instanceof Date && !isNaN(d.getTime());
}

/**
 * ============================================================================
 * NUMERIC UTILITIES
 * ============================================================================
 */

/**
 * Parse value to numeric, handling currency formatting
 * @param {*} value - Value to parse
 * @returns {number} Parsed number or 0
 */
function parseNumeric(value) {
  if (typeof value === 'number') return value;
  if (!value) return 0;
  
  const n = parseFloat(String(value).replace(/[$,\s]/g, ''));
  return isNaN(n) ? 0 : n;
}

/**
 * Format number as currency
 * @param {number} value - Number to format
 * @param {string} zeroDisplay - Display for zero (default "-")
 * @returns {string} Formatted currency string
 */
function formatCurrency(value, zeroDisplay = "-") {
  if (value === 0 || value === null || value === undefined) return zeroDisplay;
  return '$' + value.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

/**
 * Format number with specific decimals
 * @param {number} value - Number to format
 * @param {number} decimals - Decimal places
 * @returns {string} Formatted number string
 */
function formatNumber(value, decimals = 2) {
  if (value === null || value === undefined) return "";
  return Number(value).toFixed(decimals);
}

/**
 * ============================================================================
 * STRING UTILITIES
 * ============================================================================
 */

/**
 * Clean string value (trim, handle nulls)
 * @param {*} value - Value to clean
 * @returns {string} Cleaned string
 */
function cleanString(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

/**
 * Check if string is valid email
 * @param {string} email - Email to validate
 * @returns {boolean} True if valid email format
 */
function isValidEmail(email) {
  if (!email) return false;
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email.toString().trim());
}

/**
 * Normalize string for comparison
 * @param {string} value - Value to normalize
 * @returns {string} Lowercase trimmed string
 */
function normalizeString(value) {
  if (!value) return "";
  return String(value).trim().toLowerCase();
}

/**
 * Extract domain from email
 * @param {string} email - Email address
 * @returns {string} Domain or empty string
 */
function extractDomain(email) {
  if (!email || !email.includes('@')) return "";
  return email.split('@')[1].toLowerCase();
}

/**
 * ============================================================================
 * ARRAY & OBJECT UTILITIES
 * ============================================================================
 */

/**
 * Group array of objects by key
 * @param {Array} arr - Array to group
 * @param {string|Function} key - Key to group by
 * @returns {Object} Grouped object
 */
function groupBy(arr, key) {
  if (!arr || !Array.isArray(arr)) return {};
  
  return arr.reduce((result, item) => {
    const groupKey = typeof key === 'function' ? key(item) : item[key];
    if (!result[groupKey]) result[groupKey] = [];
    result[groupKey].push(item);
    return result;
  }, {});
}

/**
 * Sort array of objects by key
 * @param {Array} arr - Array to sort
 * @param {string} key - Key to sort by
 * @param {boolean} ascending - Sort direction (default true)
 * @returns {Array} Sorted array
 */
function sortBy(arr, key, ascending = true) {
  if (!arr || !Array.isArray(arr)) return [];
  
  return [...arr].sort((a, b) => {
    const aVal = typeof key === 'function' ? key(a) : a[key];
    const bVal = typeof key === 'function' ? key(b) : b[key];
    
    if (aVal < bVal) return ascending ? -1 : 1;
    if (aVal > bVal) return ascending ? 1 : -1;
    return 0;
  });
}

/**
 * Unique array by key
 * @param {Array} arr - Array to deduplicate
 * @param {string} key - Key to check uniqueness
 * @returns {Array} Deduplicated array
 */
function uniqueBy(arr, key) {
  if (!arr || !Array.isArray(arr)) return [];
  
  const seen = new Set();
  return arr.filter(item => {
    const val = typeof key === 'function' ? key(item) : item[key];
    if (seen.has(val)) return false;
    seen.add(val);
    return true;
  });
}

/**
 * Flatten nested arrays
 * @param {Array} arr - Array to flatten
 * @returns {Array} Flattened array
 */
function flatten(arr) {
  if (!arr || !Array.isArray(arr)) return [];
  return arr.reduce((flat, item) => {
    return flat.concat(Array.isArray(item) ? flatten(item) : item);
  }, []);
}

/**
 * ============================================================================
 * VALIDATION UTILITIES
 * ============================================================================
 */

/**
 * Check if value is non-empty string
 * @param {*} value - Value to check
 * @returns {boolean} True if non-empty string
 */
function isNonEmptyString(value) {
  return typeof value === 'string' && value.trim().length > 0;
}

/**
 * Check if value is numeric and greater than zero
 * @param {*} value - Value to check
 * @returns {boolean} True if positive number
 */
function isPositiveNumber(value) {
  const num = parseNumeric(value);
  return num > 0;
}

/**
 * Check if value is in array
 * @param {*} value - Value to check
 * @param {Array} arr - Array to check against
 * @returns {boolean} True if value in array
 */
function isInArray(value, arr) {
  if (!arr || !Array.isArray(arr)) return false;
  return arr.includes(value);
}

/**
 * ============================================================================
 * SHEET UTILITIES
 * ============================================================================
 */

/**
 * Get or create a sheet
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @param {string} sheetName - Name of sheet
 * @param {boolean} hidden - Whether to hide sheet
 * @returns {Sheet} Sheet instance
 */
function getOrCreateSheet(ss, sheetName, hidden = false) {
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (hidden) {
      sheet.hideSheet();
    }
    logInfo(`Created new sheet: ${sheetName}`);
  }
  
  return sheet;
}

/**
 * Safe get range values with bounds checking
 * @param {Sheet} sheet - Sheet instance
 * @param {number} row - Starting row
 * @param {number} col - Starting column
 * @param {number} numRows - Number of rows
 * @param {number} numCols - Number of columns
 * @returns {Array} Range values
 */
function safeGetRangeValues(sheet, row, col, numRows, numCols) {
  try {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    // Adjust bounds if needed
    if (row > lastRow) return [];
    if (col > lastCol) return [];
    
    const actualRows = Math.min(numRows, lastRow - row + 1);
    const actualCols = Math.min(numCols, lastCol - col + 1);
    
    if (actualRows <= 0 || actualCols <= 0) return [];
    
    return sheet.getRange(row, col, actualRows, actualCols).getValues();
  } catch (error) {
    logError(`Error getting range values: ${error.message}`);
    return [];
  }
}

/**
 * Batch write values to sheet for performance
 * @param {Sheet} sheet - Sheet instance
 * @param {number} startRow - Starting row
 * @param {number} startCol - Starting column
 * @param {Array} data - Data to write
 * @param {number} batchSize - Batch size for writing
 */
function batchWriteToSheet(sheet, startRow, startCol, data, batchSize = 500) {
  if (!data || data.length === 0) return;
  
  const totalRows = data.length;
  const totalCols = data[0] ? data[0].length : 0;
  
  for (let i = 0; i < totalRows; i += batchSize) {
    const currentBatchSize = Math.min(batchSize, totalRows - i);
    const batchData = data.slice(i, i + currentBatchSize);
    const range = sheet.getRange(startRow + i, startCol, currentBatchSize, totalCols);
    range.setValues(batchData);
  }
}

/**
 * ============================================================================
 * MISCELLANEOUS UTILITIES
 * ============================================================================
 */

/**
 * Generate unique ID
 * @returns {string} Unique ID
 */
function generateUniqueId() {
  return Utilities.getUuid();
}

/**
 * Sleep for specified milliseconds
 * @param {number} ms - Milliseconds to sleep
 */
function sleep(ms) {
  Utilities.sleep(ms);
}

/**
 * Get month difference between two dates
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @returns {number} Month difference
 */
function getMonthDifference(startDate, endDate) {
  if (!(startDate instanceof Date) || !(endDate instanceof Date)) return 0;
  return (endDate.getFullYear() - startDate.getFullYear()) * 12 + 
         (endDate.getMonth() - startDate.getMonth());
}

/**
 * Clamp number between min and max
 * @param {number} value - Value to clamp
 * @param {number} min - Minimum value
 * @param {number} max - Maximum value
 * @returns {number} Clamped value
 */
function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}
