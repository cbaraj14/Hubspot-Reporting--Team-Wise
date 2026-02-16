/**
* ============================================================================
* UNIFIED PROGRESS LOGGER - SINGLE SOURCE OF TRUTH
* ============================================================================
* 
* Provides consistent, timestamped logging across all modules.
* 
* FEATURES:
* - Timestamped log entries with milliseconds precision
* - Multiple log levels (SUCCESS, DEBUG, INFO, WARN, ERROR)
* - Unicode symbols for quick visual identification
* - Module identification for debugging
* - Consistent formatting across all reports
* 
* USAGE:
* - Initialize: initProgressLogger('Module_Name')
* - Log levels: logSuccess(), logInfo(), logDebug(), logWarn(), logError()
* - Progress: updateProgress() with automatic timestamp
* 
* OUTPUT FORMAT:
* 2026-02-16 14:33:12.345    Module_Name    SUCCESS    ✅ Operation completed successfully
* 2026-02-16 14:33:11.123    Module_Name    DEBUG    🔍 Processing item: 42
* 
* ============================================================================
*/

/**
 * Global progress state
 */
const PROGRESS_STATE = {
  moduleName: 'System',
  startTime: null,
  logEntries: []
};

/**
 * Initialize the progress logger
 * @param {string} moduleName - Name of the module/operation
 */
function initProgressLogger(moduleName) {
  PROGRESS_STATE.moduleName = moduleName || 'System';
  PROGRESS_STATE.startTime = new Date();
  PROGRESS_STATE.logEntries = [];
  
  // Log initialization
  logInfo(`🚀 Initialized logger for module: ${PROGRESS_STATE.moduleName}`);
}

/**
 * Get current timestamp string in format: YYYY-MM-DD HH:MM:SS.mmm
 */
function getCurrentTimestamp() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const hours = String(now.getHours()).padStart(2, '0');
  const minutes = String(now.getMinutes()).padStart(2, '0');
  const seconds = String(now.getSeconds()).padStart(2, '0');
  const ms = String(now.getMilliseconds()).padStart(3, '0');
  
  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}.${ms}`;
}

/**
 * Core logging function
 * @param {string} type - Log level (SUCCESS, DEBUG, INFO, WARN, ERROR)
 * @param {string} message - Log message
 */
function logEntry(type, message) {
  const timestamp = getCurrentTimestamp();
  const module = PROGRESS_STATE.moduleName;
  const symbol = getLogSymbol(type);
  
  const logLine = {
    timestamp,
    module,
    type,
    symbol,
    message,
    formatted: `${timestamp}\t${module}\t${type}\t${symbol} ${message}`
  };
  
  PROGRESS_STATE.logEntries.push(logLine);
  
  // Also log to console
  console.log(logLine.formatted);
  
  return logLine;
}

/**
 * Get symbol for log type
 */
function getLogSymbol(type) {
  const symbols = {
    'SUCCESS': '✅',
    'DEBUG': '🔍',
    'INFO': '📝',
    'WARN': '⚠️',
    'ERROR': '❌',
    'START': '▶️',
    'END': '⏹️'
  };
  return symbols[type] || '•';
}

/**
 * Log success message
 */
function logSuccess(message) {
  return logEntry('SUCCESS', message);
}

/**
 * Log debug message
 */
function logDebug(message) {
  return logEntry('DEBUG', message);
}

/**
 * Log info message
 */
function logInfo(message) {
  return logEntry('INFO', message);
}

/**
 * Log warning message
 */
function logWarn(message) {
  return logEntry('WARN', message);
}

/**
 * Log error message
 */
function logError(message) {
  return logEntry('ERROR', message);
}

/**
 * Log operation start
 */
function logStart(operationName) {
  return logEntry('START', `Starting operation: ${operationName}`);
}

/**
 * Log operation end
 */
function logEnd(operationName, details = '') {
  const endMessage = details ? `${operationName} completed. ${details}` : `${operationName} completed`;
  return logEntry('END', endMessage);
}

/**
 * Update progress with automatic formatting
 * @param {string} module - Module name (overrides default)
 * @param {string} status - Status message
 * @param {number} current - Current count
 * @param {number} total - Total count
 */
function updateProgress(module, status, current = null, total = null) {
  let message = status;
  
  if (current !== null && total !== null) {
    const percentage = total > 0 ? Math.round((current / total) * 100) : 0;
    message = `${status} (${current}/${total} - ${percentage}%)`;
  }
  
  return logInfo(message);
}

/**
 * Progress update for sheet operations
 * @param {Sheet} sheet - Progress sheet
 * @param {string} message - Progress message
 */
function updateProgressSheet(sheet, message) {
  if (!sheet) {
    logWarn('Progress sheet not available');
    return;
  }

  ensureProgressSheetHeaders(sheet);
  
  const timestamp = getCurrentTimestamp();
  
  try {
    sheet.insertRows(2, 1);
    
    const rowValues = [[timestamp, PROGRESS_STATE.moduleName, 'INFO', message]];
    sheet.getRange(2, 1, 1, rowValues[0].length).setValues(rowValues);
    
    const maxRows = 1001;
    const lastRow = sheet.getLastRow();
    if (lastRow > maxRows) {
      sheet.deleteRows(maxRows + 1, lastRow - maxRows);
    }
    
    logInfo(`Progress updated: ${message}`);
  } catch (error) {
    logError(`Failed to update progress sheet: ${error.message}`);
  }
}

/**
 * Create or get progress sheet
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @returns {Sheet} Progress sheet
 */
function getOrCreateProgressSheet(ss) {
  let progressSheet = ss.getSheetByName('Progress Sheet');
  
  if (!progressSheet) {
    progressSheet = ss.insertSheet('Progress Sheet');
  }

  ensureProgressSheetHeaders(progressSheet);
  
  return progressSheet;
}

function ensureProgressSheetHeaders(progressSheet) {
  const headers = ['Timestamp', 'Module', 'Type', 'Progress'];
  const currentHeaders = progressSheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeaders = headers.some((header, index) => currentHeaders[index] !== header);

  if (needsHeaders) {
    progressSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    progressSheet.getRange(1, 1, 1, headers.length)
      .setBackground('#2563eb')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    progressSheet.setFrozenRows(1);
  }
}

/**
 * Print execution summary
 */
function printExecutionSummary() {
  const endTime = new Date();
  const duration = PROGRESS_STATE.startTime ? 
    Math.round((endTime - PROGRESS_STATE.startTime) / 1000 * 100) / 100 : 0;
  
  logInfo(`📊 Execution Summary for ${PROGRESS_STATE.moduleName}`);
  logInfo(`⏱️  Total Duration: ${duration} seconds`);
  logInfo(`📝 Total Log Entries: ${PROGRESS_STATE.logEntries.length}`);
  
  // Count by type
  const counts = {};
  PROGRESS_STATE.logEntries.forEach(entry => {
    counts[entry.type] = (counts[entry.type] || 0) + 1;
  });
  
  Object.keys(counts).forEach(type => {
    logInfo(`   ${type}: ${counts[type]}`);
  });
}

/**
 * Export log entries as array for writing to sheets
 * @returns {Array} Array of log entry arrays
 */
function getLogEntriesForExport() {
  return PROGRESS_STATE.logEntries.map(entry => [
    entry.timestamp,
    entry.module,
    entry.type,
    entry.symbol + ' ' + entry.message
  ]);
}

/**
 * Write logs to progress sheet
 * @param {Spreadsheet} ss - Spreadsheet instance
 */
function writeLogsToSheet(ss) {
  try {
    const progressSheet = getOrCreateProgressSheet(ss);
    const logs = getLogEntriesForExport();
    
    if (logs.length === 0) {
      logWarn('No log entries to write');
      return;
    }

    const orderedLogs = logs.slice().reverse();
    progressSheet.insertRows(2, orderedLogs.length);
    progressSheet.getRange(2, 1, orderedLogs.length, 4).setValues(orderedLogs);

    const maxRows = 1001;
    const lastRow = progressSheet.getLastRow();
    if (lastRow > maxRows) {
      progressSheet.deleteRows(maxRows + 1, lastRow - maxRows);
    }
    
    logSuccess(`Wrote ${logs.length} log entries to progress sheet`);
  } catch (error) {
    logError(`Failed to write logs to sheet: ${error.message}`);
  }
}