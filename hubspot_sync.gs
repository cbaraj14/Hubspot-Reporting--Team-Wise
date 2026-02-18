/**
 * ============================================================================
 * HUBSPOT SYNC MODULE
 * ============================================================================
 * 
 * This module handles importing deal data from HubSpot API.
 * 
 * KEY FEATURES:
 * - Syncs deals from 3 pipelines (Sales, CS, Payment)
 * - Batch processing with rate limiting (100 deals per batch, 300ms delay)
 * - Automatic deduplication by Deal ID
 * - Incremental updates based on last modified date
 * 
 * MENU FUNCTIONS:
 * - syncHubSpotDeals()          → Import new/updated deals from HubSpot
 * - updateLastModifiedForExistingDeals() → Refresh last modified dates
 * 
 * CONFIGURATION:
 * - Uses CONFIG object from Config file
 * - Requires HUBSPOT_TOKEN to be set
 * - Respects API rate limits with exponential backoff
 * 
 * DATA FLOW:
 * HubSpot API → Pipeline Sheets (Sales/CS/Payment) → Temp Sheet → Reports
 * 
 * See README.md for complete documentation.
 * ============================================================================
 */

/**
 * 1. MAINTENANCE FUNCTIONS
 * Updates 'Last Modified Date' using Deal ID.
 */
function updateLastModifiedForExistingDeals() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  Logger.log(`--- [START] Update Last Modified for Sheet: ${sheet.getName()} ---`);
  
  if (data.length <= 2) {
    Logger.log('!!! Error: Sheet has no data rows (only headers).');
    SpreadsheetApp.getUi().alert('No data found in the active sheet below the headers.');
    return;
  }

  const idIndex = 8; 
  const modifiedDateIndex = 10; 
  
  // Extraction of Deal IDs from Column I
  const dealIds = data.slice(2).map((row, index) => {
    const val = row[idIndex] ? row[idIndex].toString().trim() : '';
    const isNumeric = /^\d+$/.test(val);
    
    if (val === "") {
        // Skip empty cells silently or log if debugging
        return null;
    }
    
    if (!isNumeric) {
      Logger.log(`[VALIDATION] Row ${index + 3}: Skipping invalid ID format: "${val}"`);
      return null;
    }
    
    return val;
  }).filter(Boolean);
  
  Logger.log(`[INFO] Found ${dealIds.length} valid numeric IDs to check against HubSpot.`);

  if (dealIds.length === 0) {
    const message = `No valid numeric Deal IDs found in Column I (Row 3 onwards) on sheet "${sheet.getName()}".\n\n` +
                    `Please ensure Column I contains numeric HubSpot Deal IDs.`;
    Logger.log(`[ABORT] ${message}`);
    SpreadsheetApp.getUi().alert(message);
    return;
  }

  const updateMap = {};
  const batchSize = 100;

  try {
    for (let i = 0; i < dealIds.length; i += batchSize) {
      const chunk = dealIds.slice(i, i + batchSize);
      Logger.log(`[BATCH] Processing chunk ${Math.floor(i/batchSize) + 1} (${chunk.length} IDs)`);

      const payload = {
        filterGroups: [{
          filters: [{ propertyName: 'hs_object_id', operator: 'IN', values: chunk }]
        }],
        properties: ['hs_lastmodifieddate'],
        limit: 100
      };

      try {
        const response = urlFetchWithRetry('https://api.hubapi.com/crm/v3/objects/deals/search', {
          method: 'POST',
          headers: { 
            'Authorization': 'Bearer ' + CONFIG.HUBSPOT_TOKEN, 
            'Content-Type': 'application/json' 
          },
          payload: JSON.stringify(payload)
        });

        const results = JSON.parse(response.getContentText()).results || [];
        Logger.log(`[SUCCESS] Batch search returned ${results.length} deals.`);
        
        results.forEach(deal => {
          updateMap[deal.id] = deal.properties.hs_lastmodifieddate;
        });
      } catch (batchError) {
        Logger.log(`[RETRY] Batch search failed: ${batchError.message}. Attempting individual lookups...`);
        chunk.forEach(singleId => {
          try {
            const singleResponse = urlFetchWithRetry(`https://api.hubapi.com/crm/v3/objects/deals/${singleId}?properties=hs_lastmodifieddate`, {
              method: 'GET',
              headers: { 
                'Authorization': 'Bearer ' + CONFIG.HUBSPOT_TOKEN, 
                'Content-Type': 'application/json' 
              }
            });
            const deal = JSON.parse(singleResponse.getContentText());
            if (deal && deal.id) {
              updateMap[deal.id] = deal.properties.hs_lastmodifieddate;
              Logger.log(`[INDIVIDUAL] Updated Deal ID: ${singleId}`);
            }
          } catch (itemError) {
            Logger.log(`[ERROR] Could not fetch Deal ID ${singleId}: ${itemError.message}`);
          }
        });
      }
      if (dealIds.length > batchSize) Utilities.sleep(150);
    }

    Logger.log('[STORAGE] Writing updated timestamps to Column K...');
    const newDateValues = data.slice(2).map(row => {
      const id = row[idIndex] ? row[idIndex].toString().trim() : '';
      return [updateMap[id] || row[modifiedDateIndex]];
    });

    sheet.getRange(3, modifiedDateIndex + 1, newDateValues.length, 1).setValues(newDateValues);
    Logger.log('--- [FINISH] Update Complete ---');
    SpreadsheetApp.getUi().alert(`Last Modified Dates updated successfully.`);
  } catch (error) {
    Logger.log(`[CRITICAL] Error in update function: ${error.message}`);
    SpreadsheetApp.getUi().alert('Error updating dates: ' + error.message);
  }
}

/**
 * 2. STREAMING DEAL SYNC
 */
function syncHubSpotDeals() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  initProgressLogger('HubSpot_Import');
  const progressSheet = getOrCreateProgressSheet(ss);
  updateProgressSheet(progressSheet, 'Starting HubSpot deal sync.');
  
  try {
    Logger.log('--- [START] Syncing New/Updated Deals from HubSpot ---');
    const lastUpdated = getLastUpdatedTimestamp(ss);
    const startTimestamp = parseFlexibleDate(lastUpdated);
    Logger.log(`[QUERY] Syncing deals modified since: ${startTimestamp.toISOString()}`);
    updateProgressSheet(progressSheet, `Syncing deals modified since: ${startTimestamp.toISOString()}`);
    
    let after = null;
    let totalSynced = 0;
    let batchNumber = 1;

    do {
      Logger.log(`[FETCH] Fetching batch (Current Total: ${totalSynced})...`);
      updateProgressSheet(progressSheet, `Fetching batch ${batchNumber} (Current Total: ${totalSynced})...`);
      
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
        Logger.log(`[PROCESS] Processing ${deals.length} deals in this batch.`);
        updateProgressSheet(progressSheet, `Processing ${deals.length} deals in batch ${batchNumber}.`);
        const grouped = groupDealsByPipeline(deals);
        Object.keys(grouped).forEach(sheetName => {
          streamBatchToSheet(ss, sheetName, grouped[sheetName]);
        });
        totalSynced += deals.length;
      } else {
        updateProgressSheet(progressSheet, `No deals returned in batch ${batchNumber}.`);
      }
      after = data.paging?.next?.after || null;
      batchNumber += 1;
      Utilities.sleep(300); 
    } while (after);
    
    finalizeSheets(ss);
    Logger.log(`--- [FINISH] Stream Sync Complete. Total Deals Synced: ${totalSynced} ---`);
    updateProgressSheet(progressSheet, `Sync complete. Total Deals Synced: ${totalSynced}.`);
    SpreadsheetApp.getUi().alert('Sync Complete. Check logs for details.');
  } catch (error) {
    Logger.log(`[CRITICAL] Sync Error: ${error.message}`);
    updateProgressSheet(progressSheet, `Sync error: ${error.message}`);
    SpreadsheetApp.getUi().alert('Sync Error: ' + error.message);
  }
}

function streamBatchToSheet(ss, sheetName, deals) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`[SHEET] Creating new sheet: ${sheetName}`);
    sheet = ss.insertSheet(sheetName);
    setupSheetHeaders(sheet);
  }
  
  const existingData = sheet.getDataRange().getValues();
  const idIndex = 8;
  const existingIds = existingData.map(r => r[idIndex] ? r[idIndex].toString() : "");
  
  const newDeals = deals.filter(deal => !existingIds.includes(deal.id.toString()));
  
  if (newDeals.length > 0) {
    Logger.log(`[APPEND] Adding ${newDeals.length} new records to ${sheetName}`);
    const rows = newDeals.map(deal => dealToRow(deal, sheetName));
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
  }
}

function dealToRow(deal, sheetName) {
  const p = deal.properties;
  return [
    p.dealname || '', p.amount || '', p.createdate || '', STAGE_MAP[p.dealstage] || p.dealstage, 
    sheetName, p.hubspot_owner_id || '', '', '', deal.id, p.closedate || '', 
    p.hs_lastmodifieddate || '', '', '', '', ''
  ];
}

/**
 * 3. HELPERS & UTILITIES
 */
function finalizeSheets(ss) {
  const pipelines = Object.values(CONFIG.PIPELINE_MAP);
  const now = new Date().toLocaleString();
  pipelines.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow > 2) {
      sheet.getRange(3, 1, lastRow - 2, HEADERS.length).sort({ column: 3, ascending: false });
      applySheetFormatting(sheet);
    }
    sheet.getRange('A1').setValue('Last Updated: ' + now);
  });
}

function groupDealsByPipeline(deals) {
  return deals.reduce((acc, deal) => {
    const pipelineId = deal.properties.pipeline;
    const sheetName = CONFIG.PIPELINE_MAP[pipelineId] || CONFIG.PIPELINE_MAP['default'];
    if (!acc[sheetName]) acc[sheetName] = [];
    acc[sheetName].push(deal);
    return acc;
  }, {});
}

function getLastUpdatedTimestamp(ss) {
  const sheet = ss.getSheetByName(CONFIG.PIPELINE_MAP['default']);
  if (sheet) {
    const val = sheet.getRange('A1').getValue();
    if (val && val.toString().includes('Last Updated:')) return val.toString().replace('Last Updated: ', '').trim();
  }
  return CONFIG.DEFAULT_START;
}

function setupSheetHeaders(sheet) {
  sheet.getRange('A1').setValue('Last Updated: -').setFontStyle('italic').setFontColor('#666666');
  sheet.getRange(2, 1, 1, HEADERS.length).setValues([HEADERS]).setBackground('#2563eb').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(2);
}

function applySheetFormatting(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    const filter = sheet.getFilter();
    if (filter) filter.remove();
    sheet.getRange(2, 1, lastRow - 1, HEADERS.length).createFilter();
    for (let i = 1; i <= HEADERS.length; i++) sheet.autoResizeColumn(i);
  }
}

/**
 * 4. SHARED UTILITIES
 * These functions are shared with hubspot_enrich.gs
 */

function urlFetchWithRetry(url, options) {
  let retries = 0;
  while (retries < 4) {
    const response = UrlFetchApp.fetch(url, { ...options, muteHttpExceptions: true });
    const code = response.getResponseCode();
    if (code === 200 || code === 207) return response;
    if ([429, 503, 403].includes(code)) {
      Utilities.sleep(Math.pow(2, retries) * 1000);
      retries++;
    } else throw new Error(`API error ${code}: ${response.getContentText()}`);
  }
  return null;
}

function parseFlexibleDate(dateString) {
  if (!dateString) return new Date(0);
  let normalized = dateString.toString().trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(normalized)) normalized += 'T00:00:00.000Z';
  const parsed = Date.parse(normalized);
  return isNaN(parsed) ? new Date(0) : new Date(parsed);
}
