/**
 * ============================================================================
 * HUBSPOT ENRICH MODULE
 * ============================================================================
 * 
 * This module handles enriching deal data from HubSpot with company and
 * contact associations. Maintains local databases of contacts and companies
 * for quick lookup and manual editing.
 * 
 * KEY FEATURES:
 * - Fills missing company associations from HubSpot
 * - Fills missing contact associations from HubSpot
 * - Maintains editable Contact DB for known contacts
 * - Maintains editable Companies DB for known companies
 * - Batch processing with rate limiting
 * 
 * MENU FUNCTIONS:
 * - syncMissingCompanies()      → Fill missing company associations
 * - syncMissingContacts()       → Fill missing contact associations
 * 
 * DATABASE SHEETS:
 * - Contact DB: Contact ID, Contact Name, Contact Email, Associated Company, etc.
 * - Companies DB: Company ID, Company Name, Company Owner ID, Segment, Industry, etc.
 * 
 * CONFIGURATION:
 * - Uses CONFIG object from Config file
 * - Requires HUBSPOT_TOKEN to be set
 * - Respects API rate limits with exponential backoff
 * 
 * See README.md for complete documentation.
 * ============================================================================
 */

// ============================================================================
// 1. CONFIGURATION CONSTANTS FOR ENRICHMENT
// ============================================================================

const ENRICH_DB_HEADERS = {
  CONTACT_DB: ['Contact ID', 'Contact Name', 'Contact Email', 'Associated Company', 'Phone', 'Job Title', 'Last Updated'],
  COMPANIES_DB: ['Company ID', 'Company Name', 'Company Owner ID', 'Company Segment', 'Industry', 'Website', 'Phone', 'Last Updated']
};

// ============================================================================
// 2. CONTACT DATABASE FUNCTIONS
// ============================================================================

/**
 * Get or create the Contact DB sheet
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @returns {Sheet} Contact DB sheet
 */
function getOrCreateContactDB(ss) {
  return getOrCreateDBSheet(ss, 'Contact DB', ENRICH_DB_HEADERS.CONTACT_DB);
}

/**
 * Load contact data from Contact DB into a lookup map
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @returns {Object} Map of contactId -> contact data
 */
function loadContactDB(ss) {
  const sheet = ss.getSheetByName('Contact DB');
  if (!sheet) return {};
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return {}; // Only headers
  
  const contactMap = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const contactId = row[0] ? row[0].toString() : '';
    if (contactId) {
      contactMap[contactId] = {
        contactId: contactId,
        contactName: row[1] || '',
        contactEmail: row[2] || '',
        associatedCompany: row[3] || '',
        phone: row[4] || '',
        jobTitle: row[5] || '',
        lastUpdated: row[6] || ''
      };
    }
  }
  return contactMap;
}

/**
 * Add or update contact in Contact DB
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @param {Object} contactData - Contact data object
 */
function updateContactDB(ss, contactData) {
  const sheet = getOrCreateContactDB(ss);
  const data = sheet.getDataRange().getValues();
  
  const contactId = contactData.contactId.toString();
  const now = new Date().toISOString();
  
  // Check if contact already exists
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === contactId) {
      rowIndex = i + 1; // 1-based row index
      break;
    }
  }
  
  const rowData = [
    contactData.contactId,
    contactData.contactName || '',
    contactData.contactEmail || '',
    contactData.associatedCompany || '',
    contactData.phone || '',
    contactData.jobTitle || '',
    now
  ];
  
  if (rowIndex > 0) {
    // Update existing
    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    // Append new
    sheet.appendRow(rowData);
  }
}

/**
 * Batch update contacts in Contact DB from HubSpot data
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @param {Array} contactIds - Array of contact IDs to fetch and update
 */
function batchUpdateContactDB(ss, contactIds) {
  if (!contactIds || contactIds.length === 0) return;
  
  const uniqueIds = [...new Set(contactIds.filter(id => id && id !== 'N/A'))];
  if (uniqueIds.length === 0) return;
  
  Logger.log(`[CONTACT DB] Fetching ${uniqueIds.length} contacts from HubSpot`);
  
  // Process in batches of 100
  for (let i = 0; i < uniqueIds.length; i += 100) {
    const batch = uniqueIds.slice(i, i + 100);
    
    try {
      const payload = {
        filterGroups: [{
          filters: [{ propertyName: 'hs_object_id', operator: 'IN', values: batch }]
        }],
        properties: ['firstname', 'lastname', 'email', 'phone', 'jobtitle', 'company'],
        limit: 100
      };
      
      const response = urlFetchWithRetry('https://api.hubapi.com/crm/v3/objects/contacts/search', {
        method: 'POST',
        headers: { 
          'Authorization': 'Bearer ' + CONFIG.HUBSPOT_TOKEN, 
          'Content-Type': 'application/json' 
        },
        payload: JSON.stringify(payload)
      });
      
      const data = JSON.parse(response.getContentText());
      const results = data.results || [];
      
      results.forEach(contact => {
        const props = contact.properties;
        updateContactDB(ss, {
          contactId: contact.id,
          contactName: `${props.firstname || ''} ${props.lastname || ''}`.trim(),
          contactEmail: props.email || '',
          associatedCompany: props.company || '',
          phone: props.phone || '',
          jobTitle: props.jobtitle || ''
        });
      });
      
      Logger.log(`[CONTACT DB] Updated ${results.length} contacts`);
    } catch (error) {
      Logger.log(`[CONTACT DB ERROR] Batch update failed: ${error.message}`);
    }
    
    Utilities.sleep(300);
  }
}

// ============================================================================
// 3. COMPANY DATABASE FUNCTIONS
// ============================================================================

/**
 * Get or create the Companies DB sheet
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @returns {Sheet} Companies DB sheet
 */
function getOrCreateCompaniesDB(ss) {
  return getOrCreateDBSheet(ss, 'Companies DB', ENRICH_DB_HEADERS.COMPANIES_DB);
}

/**
 * Load company data from Companies DB into a lookup map
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @returns {Object} Map of companyId -> company data
 */
function loadCompaniesDB(ss) {
  const sheet = ss.getSheetByName('Companies DB');
  if (!sheet) return {};
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return {}; // Only headers
  
  const companyMap = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const companyId = row[0] ? row[0].toString() : '';
    if (companyId) {
      companyMap[companyId] = {
        companyId: companyId,
        companyName: row[1] || '',
        companyOwnerId: row[2] || '',
        companySegment: row[3] || '',
        industry: row[4] || '',
        website: row[5] || '',
        phone: row[6] || '',
        lastUpdated: row[7] || ''
      };
    }
  }
  return companyMap;
}

/**
 * Add or update company in Companies DB
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @param {Object} companyData - Company data object
 */
function updateCompaniesDB(ss, companyData) {
  const sheet = getOrCreateCompaniesDB(ss);
  const data = sheet.getDataRange().getValues();
  
  const companyId = companyData.companyId.toString();
  const now = new Date().toISOString();
  
  // Check if company already exists
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === companyId) {
      rowIndex = i + 1; // 1-based row index
      break;
    }
  }
  
  const rowData = [
    companyData.companyId,
    companyData.companyName || '',
    companyData.companyOwnerId || '',
    companyData.companySegment || '',
    companyData.industry || '',
    companyData.website || '',
    companyData.phone || '',
    now
  ];
  
  if (rowIndex > 0) {
    // Update existing
    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    // Append new
    sheet.appendRow(rowData);
  }
}

/**
 * Batch update companies in Companies DB from HubSpot data
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @param {Array} companyIds - Array of company IDs to fetch and update
 */
function batchUpdateCompaniesDB(ss, companyIds) {
  if (!companyIds || companyIds.length === 0) return;
  
  const uniqueIds = [...new Set(companyIds.filter(id => id && id !== 'N/A'))];
  if (uniqueIds.length === 0) return;
  
  Logger.log(`[COMPANIES DB] Fetching ${uniqueIds.length} companies from HubSpot`);
  
  // Process in batches of 100
  for (let i = 0; i < uniqueIds.length; i += 100) {
    const batch = uniqueIds.slice(i, i + 100);
    
    try {
      const payload = {
        filterGroups: [{
          filters: [{ propertyName: 'hs_object_id', operator: 'IN', values: batch }]
        }],
        properties: ['name', 'hubspot_owner_id', 'segment', 'industry', 'website', 'phone'],
        limit: 100
      };
      
      const response = urlFetchWithRetry('https://api.hubapi.com/crm/v3/objects/companies/search', {
        method: 'POST',
        headers: { 
          'Authorization': 'Bearer ' + CONFIG.HUBSPOT_TOKEN, 
          'Content-Type': 'application/json' 
        },
        payload: JSON.stringify(payload)
      });
      
      const data = JSON.parse(response.getContentText());
      const results = data.results || [];
      
      results.forEach(company => {
        const props = company.properties;
        updateCompaniesDB(ss, {
          companyId: company.id,
          companyName: props.name || '',
          companyOwnerId: props.hubspot_owner_id || '',
          companySegment: props.segment || '',
          industry: props.industry || '',
          website: props.website || '',
          phone: props.phone || ''
        });
      });
      
      Logger.log(`[COMPANIES DB] Updated ${results.length} companies`);
    } catch (error) {
      Logger.log(`[COMPANIES DB ERROR] Batch update failed: ${error.message}`);
    }
    
    Utilities.sleep(300);
  }
}

// ============================================================================
// 4. HELPER FUNCTIONS FOR DB MANAGEMENT
// ============================================================================

/**
 * Get or create a database sheet with proper headers
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @param {string} sheetName - Name of sheet
 * @param {Array} headers - Array of header names
 * @returns {Sheet} Sheet instance
 */
function getOrCreateDBSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    Logger.log(`[DB] Creating new sheet: ${sheetName}`);
    sheet = ss.insertSheet(sheetName);
    
    // Add headers
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#2563eb')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Add filter
    sheet.getRange(1, 1, 1, headers.length).createFilter();
    
    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
  }
  
  return sheet;
}

/**
 * Get enriched contact data from DB or HubSpot
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @param {string} contactId - Contact ID
 * @returns {Object} Contact data
 */
function getEnrichedContactData(ss, contactId) {
  if (!contactId || contactId === 'N/A') return null;
  
  // First check DB
  const db = loadContactDB(ss);
  if (db[contactId]) {
    return db[contactId];
  }
  
  // Not in DB, fetch from HubSpot and add to DB
  try {
    const response = urlFetchWithRetry(`https://api.hubapi.com/crm/v3/objects/contacts/${contactId}?properties=firstname,lastname,email,phone,jobtitle,company`, {
      method: 'GET',
      headers: { 
        'Authorization': 'Bearer ' + CONFIG.HUBSPOT_TOKEN, 
        'Content-Type': 'application/json' 
      }
    });
    
    const contact = JSON.parse(response.getContentText());
    const props = contact.properties;
    
    const contactData = {
      contactId: contactId,
      contactName: `${props.firstname || ''} ${props.lastname || ''}`.trim(),
      contactEmail: props.email || '',
      associatedCompany: props.company || '',
      phone: props.phone || '',
      jobTitle: props.jobtitle || ''
    };
    
    updateContactDB(ss, contactData);
    return contactData;
    
  } catch (error) {
    Logger.log(`[ENRICH ERROR] Failed to fetch contact ${contactId}: ${error.message}`);
    return null;
  }
}

/**
 * Get enriched company data from DB or HubSpot
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @param {string} companyId - Company ID
 * @returns {Object} Company data
 */
function getEnrichedCompanyData(ss, companyId) {
  if (!companyId || companyId === 'N/A') return null;
  
  // First check DB
  const db = loadCompaniesDB(ss);
  if (db[companyId]) {
    return db[companyId];
  }
  
  // Not in DB, fetch from HubSpot and add to DB
  try {
    const response = urlFetchWithRetry(`https://api.hubapi.com/crm/v3/objects/companies/${companyId}?properties=name,hubspot_owner_id,segment,industry,website,phone`, {
      method: 'GET',
      headers: { 
        'Authorization': 'Bearer ' + CONFIG.HUBSPOT_TOKEN, 
        'Content-Type': 'application/json' 
      }
    });
    
    const company = JSON.parse(response.getContentText());
    const props = company.properties;
    
    const companyData = {
      companyId: companyId,
      companyName: props.name || '',
      companyOwnerId: props.hubspot_owner_id || '',
      companySegment: props.segment || '',
      industry: props.industry || '',
      website: props.website || '',
      phone: props.phone || ''
    };
    
    updateCompaniesDB(ss, companyData);
    return companyData;
    
  } catch (error) {
    Logger.log(`[ENRICH ERROR] Failed to fetch company ${companyId}: ${error.message}`);
    return null;
  }
}

// ============================================================================
// 5. ASSOCIATION SYNC FUNCTIONS
// ============================================================================

/**
 * Sync missing company associations and enrich with DB data
 */
function syncMissingCompanies() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Ensure DB sheets exist
  getOrCreateContactDB(ss);
  getOrCreateCompaniesDB(ss);
  
  const pipelines = Object.values(CONFIG.PIPELINE_MAP);
  let totalUpdated = 0;

  Logger.log('--- [START] Company Association Sync ---');

  // Collect all company IDs to update DB at end
  const allCompanyIds = [];

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
      Logger.log(`[INFO] Found ${missing.length} deals missing company info in ${sheetName}`);
      for (let i = 0; i < missing.length; i += 100) {
        const batch = missing.slice(i, i + 100);
        const batchIds = batch.map(item => item.id);
        const associations = callHubSpotBatchAssociations(batchIds, 'companies');
        
        batch.forEach((item) => {
          const foundId = associations[item.id];
          if (foundId) {
            sheet.getRange(item.row, COMPANY_ID_COL + 1).setValue(foundId);
            allCompanyIds.push(foundId);
            
            // Enrich with company name from DB/HubSpot
            const companyData = getEnrichedCompanyData(ss, foundId);
            if (companyData && companyData.companyName) {
              sheet.getRange(item.row, COMPANY_NAME_COL + 1).setValue(companyData.companyName);
            }
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
  
  // Update Companies DB with all fetched companies
  if (allCompanyIds.length > 0) {
    Logger.log(`[COMPANIES DB] Enriching ${allCompanyIds.length} companies`);
    batchUpdateCompaniesDB(ss, allCompanyIds);
  }
  
  Logger.log(`--- [FINISH] Company Sync. Total Checked: ${totalUpdated} ---`);
}

/**
 * Sync missing contact associations and enrich with DB data
 */
function syncMissingContacts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Ensure DB sheets exist
  getOrCreateContactDB(ss);
  getOrCreateCompaniesDB(ss);
  
  const pipelines = Object.values(CONFIG.PIPELINE_MAP);
  let totalUpdated = 0;

  Logger.log('--- [START] Contact Association Sync ---');

  // Collect all contact IDs to update DB at end
  const allContactIds = [];

  pipelines.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    const DEAL_ID_COL = 8;
    const CONTACT_ID_COL = 7;
    const CONTACT_NAME_COL = 12;
    const CONTACT_EMAIL_COL = 14;
    
    const missing = [];
    for (let i = 2; i < data.length; i++) {
      if (data[i][DEAL_ID_COL] && (!data[i][CONTACT_ID_COL] || data[i][CONTACT_ID_COL] === "")) {
        missing.push({ row: i + 1, id: data[i][DEAL_ID_COL].toString() });
      }
    }

    if (missing.length > 0) {
      Logger.log(`[INFO] Found ${missing.length} deals missing contact info in ${sheetName}`);
      for (let i = 0; i < missing.length; i += 100) {
        const batch = missing.slice(i, i + 100);
        const batchIds = batch.map(item => item.id);
        const associations = callHubSpotBatchAssociations(batchIds, 'contacts');
        
        batch.forEach((item) => {
          const contactId = associations[item.id];
          if (contactId) {
            sheet.getRange(item.row, CONTACT_ID_COL + 1).setValue(contactId);
            allContactIds.push(contactId);
            
            // Enrich with contact data from DB/HubSpot
            const contactData = getEnrichedContactData(ss, contactId);
            if (contactData) {
              if (contactData.contactName) {
                sheet.getRange(item.row, CONTACT_NAME_COL + 1).setValue(contactData.contactName);
              }
              if (contactData.contactEmail) {
                sheet.getRange(item.row, CONTACT_EMAIL_COL + 1).setValue(contactData.contactEmail);
              }
            }
          } else {
            sheet.getRange(item.row, CONTACT_ID_COL + 1).setValue("N/A");
          }
        });
        totalUpdated += batch.length;
        Utilities.sleep(300); 
      }
    }
  });
  
  // Update Contact DB with all fetched contacts
  if (allContactIds.length > 0) {
    Logger.log(`[CONTACT DB] Enriching ${allContactIds.length} contacts`);
    batchUpdateContactDB(ss, allContactIds);
  }
  
  Logger.log(`--- [FINISH] Contact Sync complete ---`);
}

/**
 * Refresh all contact data in Contact DB from HubSpot
 * Useful for manual updates or periodic refresh
 */
function refreshContactDB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Contact DB');
  
  if (!sheet) {
    Logger.log('[CONTACT DB] No Contact DB sheet found');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const contactIds = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      contactIds.push(data[i][0].toString());
    }
  }
  
  Logger.log(`[CONTACT DB] Refreshing ${contactIds.length} contacts`);
  batchUpdateContactDB(ss, contactIds);
  Logger.log('[CONTACT DB] Refresh complete');
}

/**
 * Refresh all company data in Companies DB from HubSpot
 * Useful for manual updates or periodic refresh
 */
function refreshCompaniesDB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Companies DB');
  
  if (!sheet) {
    Logger.log('[COMPANIES DB] No Companies DB sheet found');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const companyIds = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      companyIds.push(data[i][0].toString());
    }
  }
  
  Logger.log(`[COMPANIES DB] Refreshing ${companyIds.length} companies`);
  batchUpdateCompaniesDB(ss, companyIds);
  Logger.log('[COMPANIES DB] Refresh complete');
}

/**
 * Call HubSpot batch associations API
 * @param {Array} dealIds - Array of deal IDs
 * @param {string} toObjectType - 'companies' or 'contacts'
 * @returns {Object} Map of dealId -> associated object ID
 */
function callHubSpotBatchAssociations(dealIds, toObjectType) {
  const url = `https://api.hubapi.com/crm/v3/associations/deals/${toObjectType}/batch/read`;
  const payload = { inputs: dealIds.map(id => ({ id })) };
  const response = urlFetchWithRetry(url, {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + CONFIG.HUBSPOT_TOKEN, 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload)
  });
  const results = {};
  if (response) {
    const data = JSON.parse(response.getContentText());
    if (data.results) {
      data.results.forEach(res => { if (res.to && res.to.length > 0) results[res.from.id] = res.to[0].id; });
    }
  }
  return results;
}

// ============================================================================
// 6. INITIALIZATION FUNCTIONS
// ============================================================================

/**
 * Initialize enrichment system - creates DB sheets if they don't exist
 */
function initEnrichmentDBs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  getOrCreateContactDB(ss);
  getOrCreateCompaniesDB(ss);
  
  Logger.log('[INIT] Enrichment DBs initialized');
  SpreadsheetApp.getUi().alert('Enrichment DBs initialized. Contact DB and Companies DB sheets are ready.');
}
