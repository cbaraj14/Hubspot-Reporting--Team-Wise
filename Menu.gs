/**
 * ============================================================================
 * CUSTOM MENU & TRIGGERS
 * ============================================================================
 * 
 * This module creates the custom "HubSpot Sync" menu that appears in the
 * Google Sheets toolbar when the spreadsheet is opened.
 * 
 * MENU STRUCTURE:
 * 
 * HubSpot Sync
 * ├── 1. Sync New Deals              → syncHubSpotDeals()
 * ├── ─────────────────
 * ├── 2. Sync Missing Companies      → syncMissingCompanies()
 * ├── 3. Sync Missing Contacts       → syncMissingContacts()
 * ├── ─────────────────
 * ├── 4. Update ID to Names          → syncAllPipelineData()
 * ├── ─────────────────
 * ├── 5. Generate Summary            → generateSummaryReport()
 * ├── 6. Run Sales Team Report       → salesGenerateSalesTeamReport()
 * ├── 7. Run CS Team Report          → generateCSReport()
 * ├── ─────────────────
 * └── ⚙️ Auto-Refresh Settings
 *     ├── Enable Daily Auto-Refresh  → setupDailyTrigger()
 *     ├── Disable Daily Auto-Refresh → removeDailyTrigger()
 *     └── View Trigger Status        → showTriggerStatus()
 * 
 * RECOMMENDED WORKFLOW:
 * 
 * First-Time Setup:
 *   1. Sync New Deals (imports all deals from HubSpot)
 *   2. Sync Missing Companies (enriches with company associations)
 *   3. Sync Missing Contacts (enriches with contact associations)
 *   4. Update ID to Names (converts IDs to readable names)
 *   5. Generate Summary (creates comprehensive report)
 *   6. Enable Daily Auto-Refresh (for automatic updates)
 * 
 * Daily/Weekly Updates (if not using auto-refresh):
 *   1. Sync New Deals (gets only new/updated deals)
 *   2. Update ID to Names (refreshes names)
 *   3. Generate desired reports
 * 
 * AUTOMATIC TRIGGERS:
 * - onOpen() runs automatically when spreadsheet is opened
 * - Creates the custom menu
 * - No authentication required after first authorization
 * 
 * AUTO-REFRESH:
 * - setupDailyTrigger() creates a time-driven trigger for daily sync
 * - dailyAutoRefresh() runs the complete sync workflow automatically
 * - Logs all activity to AutoRefresh_Log sheet
 * - Runs at 6:00 AM daily (configurable)
 * 
 * ERROR HANDLING:
 * - Menu creation failures are logged but don't block opening
 * - Individual function errors are shown in UI alerts
 * - Auto-refresh errors are logged to AutoRefresh_Log sheet
 * 
 * See README.md for detailed usage instructions.
 * ============================================================================
 */

function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('HubSpot Sync')
      .addItem('1. Sync New Deals', 'syncHubSpotDeals')
      .addSeparator()
      .addItem('2. Sync Missing Companies', 'syncMissingCompanies')
      .addItem('3. Sync Missing Contacts', 'syncMissingContacts')
      .addSeparator()
      .addItem('4. Update ID to Names', 'syncAllPipelineData')
      .addSeparator()
      .addItem('5. Generate Summary', 'generateSummaryReport')
      .addItem('6. Run Sales Team Report', 'salesGenerateSalesTeamReport')
      .addItem('7. Run CS Team Report', 'CS_forecastRevenue')
      .addSeparator()
      .addSubMenu(ui.createMenu('⚙️ Auto-Refresh Settings')
        .addItem('Enable Daily Auto-Refresh', 'setupDailyTrigger')
        .addItem('Disable Daily Auto-Refresh', 'removeDailyTrigger')
        .addItem('View Trigger Status', 'showTriggerStatus'))
      .addToUi();
  } catch (e) {
    console.log('Menu creation skipped: ' + e.message);
  }
}

/**
 * ============================================================================
 * TIME-DRIVEN TRIGGER FUNCTIONS
 * ============================================================================
 * 
 * These functions enable automatic daily refresh of HubSpot data.
 * 
 * HOW IT WORKS:
 * 1. setupDailyTrigger() - Creates a time-driven trigger that runs every day
 * 2. dailyAutoRefresh() - The main function that runs automatically
 * 3. removeDailyTrigger() - Removes the automatic trigger
 * 4. showTriggerStatus() - Shows current trigger status
 * 
 * SETUP INSTRUCTIONS:
 * 1. Click "HubSpot Sync" → "⚙️ Auto-Refresh Settings" → "Enable Daily Auto-Refresh"
 * 2. Authorize the script when prompted
 * 3. The script will now run automatically every day at 6:00 AM
 * 
 * NOTE: The first time you enable auto-refresh, Google will ask for authorization.
 * This is required for time-driven triggers to work.
 * ============================================================================
 */

/**
 * Main function that runs automatically every day.
 * Performs the complete sync workflow.
 */
function dailyAutoRefresh() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = getOrCreateLogSheet(ss);
  const startTime = new Date();
  
  logMessage(logSheet, '========== DAILY AUTO-REFRESH STARTED ==========');
  
  try {
    // Step 1: Sync new/updated deals from HubSpot
    logMessage(logSheet, 'Step 1: Syncing deals from HubSpot...');
    syncHubSpotDealsAuto(logSheet);
    
    // Step 2: Sync missing companies
    logMessage(logSheet, 'Step 2: Syncing missing company associations...');
    syncMissingCompaniesAuto(logSheet);
    
    // Step 3: Sync missing contacts
    logMessage(logSheet, 'Step 3: Syncing missing contact associations...');
    syncMissingContactsAuto(logSheet);
    
    // Step 4: Update ID to names
    logMessage(logSheet, 'Step 4: Updating ID to names...');
    syncAllPipelineDataAuto(logSheet);
    
    const endTime = new Date();
    const duration = Math.round((endTime - startTime) / 1000);
    logMessage(logSheet, `========== DAILY AUTO-REFRESH COMPLETED in ${duration} seconds ==========`);
    
    // Send notification email (optional - uncomment if needed)
    // sendCompletionEmail(logSheet, startTime, duration);
    
  } catch (error) {
    logMessage(logSheet, `ERROR: ${error.message}`);
    logMessage(logSheet, '========== DAILY AUTO-REFRESH FAILED ==========');
    
    // Send error notification (optional - uncomment if needed)
    // sendErrorEmail(error);
  }
}

/**
 * Sets up a daily time-driven trigger at 6:00 AM.
 */
function setupDailyTrigger() {
  // Check if trigger already exists
  const triggers = ScriptApp.getProjectTriggers();
  const existingTrigger = triggers.find(t => t.getHandlerFunction() === 'dailyAutoRefresh');
  
  if (existingTrigger) {
    SpreadsheetApp.getUi().alert(
      'Daily Auto-Refresh is already enabled!\n\n' +
      'The script runs automatically every day at 6:00 AM.\n\n' +
      'To change the time, first disable auto-refresh, then re-enable it.'
    );
    return;
  }
  
  // Create a new daily trigger at 6:00 AM
  ScriptApp.newTrigger('dailyAutoRefresh')
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .create();
  
  // Create a log sheet to track auto-refresh runs
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = getOrCreateLogSheet(ss);
  logMessage(logSheet, 'Daily auto-refresh trigger created. Will run at 6:00 AM daily.');
  
  SpreadsheetApp.getUi().alert(
    '✅ Daily Auto-Refresh Enabled!\n\n' +
    'The script will now run automatically every day at 6:00 AM.\n\n' +
    'It will perform:\n' +
    '1. Sync new deals from HubSpot\n' +
    '2. Sync missing company associations\n' +
    '3. Sync missing contact associations\n' +
    '4. Update ID to names\n\n' +
    'Check the "AutoRefresh_Log" sheet for execution history.'
  );
}

/**
 * Removes the daily time-driven trigger.
 */
function removeDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const existingTrigger = triggers.find(t => t.getHandlerFunction() === 'dailyAutoRefresh');
  
  if (!existingTrigger) {
    SpreadsheetApp.getUi().alert(
      'Daily Auto-Refresh is not currently enabled.\n\n' +
      'Click "Enable Daily Auto-Refresh" to set up automatic refresh.'
    );
    return;
  }
  
  ScriptApp.deleteTrigger(existingTrigger);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = getOrCreateLogSheet(ss);
  logMessage(logSheet, 'Daily auto-refresh trigger removed by user.');
  
  SpreadsheetApp.getUi().alert(
    '✅ Daily Auto-Refresh Disabled!\n\n' +
    'The script will no longer run automatically.\n\n' +
    'You can still run manual syncs from the menu.'
  );
}

/**
 * Shows the current status of time-driven triggers.
 */
function showTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  const dailyTrigger = triggers.find(t => t.getHandlerFunction() === 'dailyAutoRefresh');
  
  let message = '';
  
  if (dailyTrigger) {
    message = 
      '📊 AUTO-REFRESH STATUS: ENABLED\n\n' +
      'Function: dailyAutoRefresh\n' +
      'Type: Time-driven (Daily)\n' +
      'Hour: 6:00 AM (server time)\n\n' +
      'The script runs automatically every day and:\n' +
      '• Syncs new deals from HubSpot\n' +
      '• Updates company associations\n' +
      '• Updates contact associations\n' +
      '• Converts IDs to names\n\n' +
      'Check "AutoRefresh_Log" sheet for execution history.';
  } else {
    message = 
      '📊 AUTO-REFRESH STATUS: DISABLED\n\n' +
      'No automatic triggers are configured.\n\n' +
      'To enable daily auto-refresh:\n' +
      '1. Click "Enable Daily Auto-Refresh" in the menu\n' +
      '2. Authorize the script when prompted\n' +
      '3. The script will run at 6:00 AM daily';
  }
  
  SpreadsheetApp.getUi().alert(message);
}

/**
 * Creates or gets the log sheet for tracking auto-refresh runs.
 */
function getOrCreateLogSheet(ss) {
  let logSheet = ss.getSheetByName('AutoRefresh_Log');
  
  if (!logSheet) {
    logSheet = ss.insertSheet('AutoRefresh_Log');
    logSheet.getRange('A1:C1').setValues([['Timestamp', 'Status', 'Message']])
      .setBackground('#2563eb')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    logSheet.setFrozenRows(1);
    logSheet.autoResizeColumns(1, 3);
  }
  
  return logSheet;
}

/**
 * Logs a message to the log sheet.
 */
function logMessage(logSheet, message) {
  const timestamp = new Date();
  const status = message.includes('ERROR') || message.includes('FAILED') ? '❌ Error' : 
                 message.includes('COMPLETED') ? '✅ Success' : 
                 message.includes('STARTED') ? '🔄 Started' : 'ℹ️ Info';
  
  logSheet.appendRow([timestamp, status, message]);
  
  // Keep only last 1000 rows to prevent sheet from growing too large
  const lastRow = logSheet.getLastRow();
  if (lastRow > 1000) {
    logSheet.deleteRows(2, lastRow - 1000);
  }
  
  console.log(`[${timestamp.toISOString()}] ${message}`);
}

/**
 * Silent version of syncHubSpotDeals for auto-refresh (no UI alerts).
 */
function syncHubSpotDealsAuto(logSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const lastUpdated = getLastUpdatedTimestamp(ss);
    const startTimestamp = parseFlexibleDate(lastUpdated);
    logMessage(logSheet, `Syncing deals modified since: ${startTimestamp.toISOString()}`);
    
    let after = null;
    let totalSynced = 0;

    do {
      const payload = {
        filterGroups: [{
          filters: [
            { propertyName: 'hs_lastmodifieddate', operator: 'GTE', value: startTimestamp.getTime().toString() },
            { propertyName: 'dealstage', operator: 'IN', values: STAGE_IDS }
          ]
        }],
        properties: ['dealname', 'dealstage', 'pipeline', 'amount', 'closedate', 'createdate', 'hubspot_owner_id', 'hs_lastmodifieddate'],
        limit: 100,
        after: after
      };

      const response = urlFetchWithRetry('https://api.hubapi.com/crm/v3/objects/deals/search', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + CONFIG.HUBSPOT_TOKEN, 'Content-Type': 'application/json' },
        payload: JSON.stringify(payload)
      });

      const data = JSON.parse(response.getContentText());
      const deals = data.results || [];
      
      if (deals.length > 0) {
        const grouped = groupDealsByPipeline(deals);
        Object.keys(grouped).forEach(sheetName => {
          streamBatchToSheet(ss, sheetName, grouped[sheetName]);
        });
        totalSynced += deals.length;
      }
      after = data.paging?.next?.after || null;
      Utilities.sleep(300); 
    } while (after);
    
    finalizeSheets(ss);
    logMessage(logSheet, `Deal sync complete. Total deals synced: ${totalSynced}`);
    
  } catch (error) {
    logMessage(logSheet, `ERROR in deal sync: ${error.message}`);
    throw error;
  }
}

/**
 * Silent version of syncMissingCompanies for auto-refresh.
 */
function syncMissingCompaniesAuto(logSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pipelines = Object.values(CONFIG.PIPELINE_MAP);
  let totalUpdated = 0;

  pipelines.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    const DEAL_ID_COL = 8;
    const COMPANY_ID_COL = 6;
    const COMPANY_NAME_COL = 11;
    const CONTACT_NAME_COL = 12;
    const CONTACT_EMAIL_COL = 14;
    
    const missing = [];
    for (let i = 2; i < data.length; i++) {
      if (data[i][DEAL_ID_COL] && (!data[i][COMPANY_ID_COL] || data[i][COMPANY_ID_COL] === "")) {
        missing.push({ 
          row: i + 1, 
          id: data[i][DEAL_ID_COL].toString(),
          contactName: data[i][CONTACT_NAME_COL],
          contactEmail: data[i][CONTACT_EMAIL_COL]
        });
      }
    }

    if (missing.length > 0) {
      for (let i = 0; i < missing.length; i += 100) {
        const batch = missing.slice(i, i + 100);
        const batchIds = batch.map(item => item.id);
        const associations = callHubSpotBatchAssociations(batchIds, 'companies');
        
        batch.forEach((item) => {
          const foundId = associations[item.id];
          if (foundId) {
            sheet.getRange(item.row, COMPANY_ID_COL + 1).setValue(foundId);
          } else {
            let fallbackName = "";
            if (item.contactName && item.contactName !== "N/A" && item.contactName !== "") {
              fallbackName = `Individual-(${item.contactName})`;
            } else if (item.contactEmail && item.contactEmail !== "") {
              fallbackName = item.contactEmail;
            }
            if (fallbackName !== "") sheet.getRange(item.row, COMPANY_NAME_COL + 1).setValue(fallbackName);
          }
        });
        totalUpdated += batch.length;
        Utilities.sleep(300);
      }
    }
  });
  
  logMessage(logSheet, `Company sync complete. Total checked: ${totalUpdated}`);
}

/**
 * Silent version of syncMissingContacts for auto-refresh.
 */
function syncMissingContactsAuto(logSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pipelines = Object.values(CONFIG.PIPELINE_MAP);
  let totalUpdated = 0;

  pipelines.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    const DEAL_ID_COL = 8;
    const CONTACT_ID_COL = 7;
    
    const missing = [];
    for (let i = 2; i < data.length; i++) {
      if (data[i][DEAL_ID_COL] && (!data[i][CONTACT_ID_COL] || data[i][CONTACT_ID_COL] === "")) {
        missing.push({ row: i + 1, id: data[i][DEAL_ID_COL].toString() });
      }
    }

    if (missing.length > 0) {
      for (let i = 0; i < missing.length; i += 100) {
        const batch = missing.slice(i, i + 100);
        const batchIds = batch.map(item => item.id);
        const associations = callHubSpotBatchAssociations(batchIds, 'contacts');
        
        batch.forEach((item) => {
          sheet.getRange(item.row, CONTACT_ID_COL + 1).setValue(associations[item.id] || "N/A");
        });
        totalUpdated += batch.length;
        Utilities.sleep(300);
      }
    }
  });
  
  logMessage(logSheet, `Contact sync complete. Total checked: ${totalUpdated}`);
}

/**
 * Silent version of syncAllPipelineData for auto-refresh.
 */
function syncAllPipelineDataAuto(logSheet) {
  try {
    syncAllPipelineData();
    logMessage(logSheet, 'ID to names update complete.');
  } catch (error) {
    logMessage(logSheet, `ERROR in ID to names update: ${error.message}`);
    throw error;
  }
}

/**
 * Optional: Send email notification on completion.
 * Uncomment and configure the email address to enable.
 */
function sendCompletionEmail(logSheet, startTime, duration) {
  const recipient = 'your-email@example.com'; // Configure this
  const subject = 'HubSpot Daily Sync Completed Successfully';
  const body = `
The daily HubSpot sync completed successfully.

Start Time: ${startTime.toLocaleString()}
Duration: ${duration} seconds

Check the AutoRefresh_Log sheet in your spreadsheet for detailed logs.
  `;
  
  MailApp.sendEmail(recipient, subject, body);
}

/**
 * Optional: Send email notification on error.
 * Uncomment and configure the email address to enable.
 */
function sendErrorEmail(error) {
  const recipient = 'your-email@example.com'; // Configure this
  const subject = 'HubSpot Daily Sync FAILED - Action Required';
  const body = `
The daily HubSpot sync encountered an error.

Error: ${error.message}

Please check the AutoRefresh_Log sheet and resolve the issue.
  `;
  
  MailApp.sendEmail(recipient, subject, body);
}
