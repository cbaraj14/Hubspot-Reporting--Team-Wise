/**
 * ============================================================================
 * HUBSPOT ENRICH MODULE
 * ============================================================================
 * 
 * This module handles enriching deal data from HubSpot with company,
 * contact, and owner associations. Maintains local databases for quick
 * lookup and manual editing.
 * 
 * KEY FEATURES:
 * - Fills missing company associations from HubSpot
 * - Fills missing contact associations from HubSpot
 * - Fills missing owner information from Owner DB or HubSpot
 * - Maintains editable Contact DB for known contacts
 * - Maintains editable Companies DB for known companies
 * - Maintains editable Owner DB for known owners
 * - Batch processing with rate limiting
 * 
 * MENU FUNCTIONS:
 * - syncMissingCompanies()      → Fill missing company associations
 * - syncMissingContacts()       → Fill missing contact associations
 * - syncMissingOwners()         → Fill missing owner names
 * 
 * DATABASE SHEETS:
 * - Contact DB: Contact ID, Contact Name, Contact Email, Associated Company, etc.
 * - Companies DB: Company ID, Company Name, Company Owner ID, Segment, Industry, etc.
 * - Owner DB: Owner ID, Owner Name, Team Name
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
  COMPANIES_DB: ['Company ID', 'Company Name', 'Company Owner ID', 'Company Segment', 'Industry', 'Website', 'Phone', 'Last Updated'],
  OWNER_DB: ['Owner ID', 'Owner Name', 'Team Name', 'Last Updated']
};

// ============================================================================
// 2. OWNER DATABASE FUNCTIONS
// ============================================================================

/**
 * Get or create the Owner DB sheet
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @returns {Sheet} Owner DB sheet
 */
function getOrCreateOwnerDB(ss) {
  return getOrCreateDBSheet(ss, 'Owner DB', ENRICH_DB_HEADERS.OWNER_DB);
}

/**
 * Load owner data from Owner DB into a lookup map
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @returns {Object} Map of ownerId -> owner data
 */
function loadOwnerDB(ss) {
  const sheet = ss.getSheetByName('Owner DB');
  if (!sheet) return {};
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return {}; // Only headers
  
  const ownerMap = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const ownerId = row[0] ? row[0].toString() : '';
    if (ownerId) {
      ownerMap[ownerId] = {
        ownerId: ownerId,
        ownerName: row[1] || '',
        teamName: row[2] || '',
        lastUpdated: row[3] || ''
      };
    }
  }
  return ownerMap;
}

/**
 * Add or update owner in Owner DB
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @param {Object} ownerData - Owner data object
 */
function updateOwnerDB(ss, ownerData) {
  const sheet = getOrCreateOwnerDB(ss);
  const data = sheet.getDataRange().getValues();
  
  const ownerId = ownerData.ownerId.toString();
  const now = new Date().toISOString();
  
  // Check if owner already exists
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === ownerId) {
      rowIndex = i + 1; // 1-based row index
      break;
    }
  }
  
  const rowData = [
    ownerData.ownerId,
    ownerData.ownerName || '',
    ownerData.teamName || '',
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
 * Fetch owner data from HubSpot API
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @param {string} ownerId - Owner ID
 * @returns {Object} Owner data or null
 */
function fetchOwnerFromHubSpot(ss, ownerId) {
  if (!ownerId || ownerId === 'N/A' || ownerId === '') return null;
  
  try {
    const response = urlFetchWithRetry(`https://api.hubapi.com/crm/v3/owners/${ownerId}`, {
      method: 'GET',
      headers: { 
        'Authorization': 'Bearer ' + CONFIG.HUBSPOT_TOKEN, 
        'Content-Type': 'application/json' 
      }
    });
    
    const owner = JSON.parse(response.getContentText());
    
    const ownerData = {
      ownerId: ownerId,
      ownerName: `${owner.firstName || ''} ${owner.lastName || ''}`.trim(),
      teamName: owner.teamName || ''
    };
    
    // Add to DB
    updateOwnerDB(ss, ownerData);
    
    return ownerData;
    
  } catch (error) {
    Logger.log(`[OWNER FETCH ERROR] Failed to fetch owner ${ownerId}: ${error.message}`);
    return null;
  }
}

/**
 * Get owner data from DB or HubSpot
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @param {string} ownerId - Owner ID
 * @returns {Object} Owner data or null
 */
function getEnrichedOwnerData(ss, ownerId) {
  if (!ownerId || ownerId === 'N/A' || ownerId === '') return null;
  
  // First check DB
  const db = loadOwnerDB(ss);
  if (db[ownerId]) {
    return db[ownerId];
  }
  
  // Not in DB, fetch from HubSpot
  return fetchOwnerFromHubSpot(ss, ownerId);
}

// ============================================================================
// 3. CONTACT DATABASE FUNCTIONS
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
 * Collect unique contact IDs from all pipeline sheets
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @returns {Array} Array of unique contact IDs
 */
function collectContactIdsFromPipelines(ss) {
  const pipelines = Object.values(CONFIG.PIPELINE_MAP);
  const contactIds = new Set();
  
  pipelines.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    for (let i = 2; i < data.length; i++) {
      const contactId = data[i][7]; // Column H - Contact ID
      if (contactId && contactId !== '' && contactId !== 'N/A') {
        contactIds.add(contactId.toString());
      }
    }
  });
  
  return Array.from(contactIds);
}

/**
 * Collect unique company IDs from all pipeline sheets
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @returns {Array} Array of unique company IDs
 */
function collectCompanyIdsFromPipelines(ss) {
  const pipelines = Object.values(CONFIG.PIPELINE_MAP);
  const companyIds = new Set();
  
  pipelines.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    for (let i = 2; i < data.length; i++) {
      const companyId = data[i][6]; // Column G - Company ID
      if (companyId && companyId !== '' && companyId !== 'N/A') {
        companyIds.add(companyId.toString());
      }
    }
  });
  
  return Array.from(companyIds);
}

/**
 * Collect unique owner IDs from all pipeline sheets
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @returns {Array} Array of unique owner IDs
 */
function collectOwnerIdsFromPipelines(ss) {
  const pipelines = Object.values(CONFIG.PIPELINE_MAP);
  const ownerIds = new Set();
  
  pipelines.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    for (let i = 2; i < data.length; i++) {
      const ownerId = data[i][5]; // Column F - Owner ID
      if (ownerId && ownerId !== '' && ownerId !== 'N/A') {
        ownerIds.add(ownerId.toString());
      }
    }
  });
  
  return Array.from(ownerIds);
}

/**
 * Fetch contacts by IDs from HubSpot API (batch optimized)
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @param {Array} contactIds - Array of contact IDs to fetch
 */
function fetchContactsByIds(ss, contactIds) {
  if (!contactIds || contactIds.length === 0) {
    Logger.log('[CONTACT DB] No contact IDs to fetch');
    return;
  }
  
  Logger.log(`[CONTACT DB] Fetching ${contactIds.length} contacts from HubSpot...`);
  
  const existingDB = loadContactDB(ss);
  const idsToFetch = contactIds.filter(id => !existingDB[id]);
  
  if (idsToFetch.length === 0) {
    Logger.log('[CONTACT DB] All contacts already in DB');
    return;
  }
  
  Logger.log(`[CONTACT DB] ${idsToFetch.length} new contacts to fetch`);
  
  // Process in batches of 100 using search API
  for (let i = 0; i < idsToFetch.length; i += 100) {
    const batch = idsToFetch.slice(i, i + 100);
    
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
      
      // Update DB with fetched contacts
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
      
      Logger.log(`[CONTACT DB] Fetched ${results.length}/${batch.length} contacts in batch`);
    } catch (error) {
      Logger.log(`[CONTACT DB ERROR] Batch fetch failed: ${error.message}`);
      
      // Fall back to individual fetches for this batch
      batch.forEach(contactId => {
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
          updateContactDB(ss, {
            contactId: contactId,
            contactName: `${props.firstname || ''} ${props.lastname || ''}`.trim(),
            contactEmail: props.email || '',
            associatedCompany: props.company || '',
            phone: props.phone || '',
            jobTitle: props.jobtitle || ''
          });
        } catch (individualError) {
          Logger.log(`[CONTACT DB ERROR] Failed to fetch contact ${contactId}: ${individualError.message}`);
        }
        Utilities.sleep(100);
      });
    }
    
    Utilities.sleep(300);
  }
  
  Logger.log(`[CONTACT DB] Fetch complete`);
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

// ============================================================================
// 4. COMPANY DATABASE FUNCTIONS
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
 * Fetch companies by IDs from HubSpot API (batch optimized)
 * @param {Spreadsheet} ss - Spreadsheet instance
 * @param {Array} companyIds - Array of company IDs to fetch
 */
function fetchCompaniesByIds(ss, companyIds) {
  if (!companyIds || companyIds.length === 0) {
    Logger.log('[COMPANIES DB] No company IDs to fetch');
    return;
  }
  
  Logger.log(`[COMPANIES DB] Fetching ${companyIds.length} companies from HubSpot...`);
  
  const existingDB = loadCompaniesDB(ss);
  const idsToFetch = companyIds.filter(id => !existingDB[id]);
  
  if (idsToFetch.length === 0) {
    Logger.log('[COMPANIES DB] All companies already in DB');
    return;
  }
  
  Logger.log(`[COMPANIES DB] ${idsToFetch.length} new companies to fetch`);
  
  // Process in batches of 100 using search API
  for (let i = 0; i < idsToFetch.length; i += 100) {
    const batch = idsToFetch.slice(i, i + 100);
    
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
      
      // Update DB with fetched companies
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
      
      Logger.log(`[COMPANIES DB] Fetched ${results.length}/${batch.length} companies in batch`);
    } catch (error) {
      Logger.log(`[COMPANIES DB ERROR] Batch fetch failed: ${error.message}`);
      
      // Fall back to individual fetches for this batch
      batch.forEach(companyId => {
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
          updateCompaniesDB(ss, {
            companyId: companyId,
            companyName: props.name || '',
            companyOwnerId: props.hubspot_owner_id || '',
            companySegment: props.segment || '',
            industry: props.industry || '',
            website: props.website || '',
            phone: props.phone || ''
          });
        } catch (individualError) {
          Logger.log(`[COMPANIES DB ERROR] Failed to fetch company ${companyId}: ${individualError.message}`);
        }
        Utilities.sleep(100);
      });
    }
    
    Utilities.sleep(300);
  }
  
  Logger.log(`[COMPANIES DB] Fetch complete`);
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
// 5. HELPER FUNCTIONS FOR DB MANAGEMENT
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
 * Initialize all enrichment DBs
 * Creates sheets if they don't exist
 * Collects IDs from pipeline sheets and fetches only those records
 * @param {Spreadsheet} ss - Spreadsheet instance
 */
function initializeEnrichmentDBs(ss) {
  Logger.log('--- [INIT] Initializing Enrichment DBs ---');
  
  // Create/get DB sheets
  getOrCreateContactDB(ss);
  getOrCreateCompaniesDB(ss);
  getOrCreateOwnerDB(ss);
  
  // Collect IDs from pipeline sheets
  Logger.log('[INIT] Collecting IDs from pipeline sheets...');
  const contactIds = collectContactIdsFromPipelines(ss);
  const companyIds = collectCompanyIdsFromPipelines(ss);
  const ownerIds = collectOwnerIdsFromPipelines(ss);
  
  Logger.log(`[INIT] Found in pipelines: ${contactIds.length} contacts, ${companyIds.length} companies, ${ownerIds.length} owners`);
  
  // Fetch only the records that are referenced in pipelines
  if (contactIds.length > 0) {
    fetchContactsByIds(ss, contactIds);
  }
  if (companyIds.length > 0) {
    fetchCompaniesByIds(ss, companyIds);
  }
  if (ownerIds.length > 0) {
    // Fetch owners individually (no batch search available)
    const existingOwners = loadOwnerDB(ss);
    ownerIds.forEach(ownerId => {
      if (!existingOwners[ownerId]) {
        fetchOwnerFromHubSpot(ss, ownerId);
      }
    });
  }
  
  Logger.log('--- [INIT] Enrichment DBs initialized ---');
}

// ============================================================================
// 6. ASSOCIATION SYNC FUNCTIONS
// ============================================================================

/**
 * Sync missing company associations and enrich with DB data
 */
function syncMissingCompanies() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Initialize DBs (creates if not exists, fetches all if new)
  initializeEnrichmentDBs(ss);
  
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
  
  Logger.log(`--- [FINISH] Company Sync. Total Checked: ${totalUpdated} ---`);
}

/**
 * Sync missing contact associations and enrich with DB data
 */
function syncMissingContacts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Initialize DBs (creates if not exists, fetches all if new)
  initializeEnrichmentDBs(ss);
  
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
  
  Logger.log(`--- [FINISH] Contact Sync complete ---`);
}

/**
 * Sync missing owner information and enrich with Owner DB data
 */
function syncMissingOwners() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Ensure Owner DB exists
  getOrCreateOwnerDB(ss);
  
  const pipelines = Object.values(CONFIG.PIPELINE_MAP);
  let totalUpdated = 0;

  Logger.log('--- [START] Owner Sync ---');

  pipelines.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    const OWNER_ID_COL = 5;  // Column F - Owner ID
    const OWNER_NAME_COL = 13; // Column N - Owner Name
    
    const missingOwners = [];
    for (let i = 2; i < data.length; i++) {
      const ownerId = data[i][OWNER_ID_COL];
      const ownerName = data[i][OWNER_NAME_COL];
      
      // If we have an owner ID but no owner name
      if (ownerId && ownerId !== '' && (!ownerName || ownerName === '')) {
        missingOwners.push({ 
          row: i + 1, 
          ownerId: ownerId.toString()
        });
      }
    }

    if (missingOwners.length > 0) {
      Logger.log(`[INFO] Found ${missingOwners.length} deals missing owner info in ${sheetName}`);
      
      // Process unique owner IDs to avoid duplicate API calls
      const uniqueOwnerIds = [...new Set(missingOwners.map(item => item.ownerId))];
      
      uniqueOwnerIds.forEach(ownerId => {
        const ownerData = getEnrichedOwnerData(ss, ownerId);
        
        // Update all rows with this owner ID
        missingOwners
          .filter(item => item.ownerId === ownerId)
          .forEach(item => {
            if (ownerData && ownerData.ownerName) {
              sheet.getRange(item.row, OWNER_NAME_COL + 1).setValue(ownerData.ownerName);
            }
          });
        
        totalUpdated++;
        Utilities.sleep(100); // Rate limiting
      });
    }
  });
  
  Logger.log(`--- [FINISH] Owner Sync. Total Updated: ${totalUpdated} ---`);
}

/**
 * Scan pipelines for new contact/company/owner IDs and add to DBs
 * This should be called periodically or after deal sync
 */
function syncNewIDsToDBs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  Logger.log('--- [START] Syncing New IDs to DBs ---');
  
  const pipelines = Object.values(CONFIG.PIPELINE_MAP);
  const newContactIds = new Set();
  const newCompanyIds = new Set();
  const newOwnerIds = new Set();
  
  // Load existing DBs
  const contactDB = loadContactDB(ss);
  const companyDB = loadCompaniesDB(ss);
  const ownerDB = loadOwnerDB(ss);
  
  pipelines.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    
    for (let i = 2; i < data.length; i++) {
      const contactId = data[i][7];  // Column H - Contact ID
      const companyId = data[i][6];  // Column G - Company ID
      const ownerId = data[i][5];    // Column F - Owner ID
      
      if (contactId && contactId !== '' && contactId !== 'N/A' && !contactDB[contactId]) {
        newContactIds.add(contactId.toString());
      }
      if (companyId && companyId !== '' && companyId !== 'N/A' && !companyDB[companyId]) {
        newCompanyIds.add(companyId.toString());
      }
      if (ownerId && ownerId !== '' && ownerId !== 'N/A' && !ownerDB[ownerId]) {
        newOwnerIds.add(ownerId.toString());
      }
    }
  });
  
  Logger.log(`[SCAN] Found ${newContactIds.size} new contacts, ${newCompanyIds.size} new companies, ${newOwnerIds.size} new owners`);
  
  // Fetch and add new contacts using batch API
  if (newContactIds.size > 0) {
    fetchContactsByIds(ss, Array.from(newContactIds));
  }
  
  // Fetch and add new companies using batch API
  if (newCompanyIds.size > 0) {
    fetchCompaniesByIds(ss, Array.from(newCompanyIds));
  }
  
  // Fetch and add new owners (individual API calls - no batch search available)
  if (newOwnerIds.size > 0) {
    Logger.log(`[OWNER DB] Adding ${newOwnerIds.size} new owners`);
    newOwnerIds.forEach(ownerId => {
      fetchOwnerFromHubSpot(ss, ownerId);
      Utilities.sleep(100);
    });
  }
  
  Logger.log('--- [FINISH] New IDs synced to DBs ---');
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
// 7. REFRESH FUNCTIONS
// ============================================================================

/**
 * Refresh all contact data in Contact DB from HubSpot
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
  
  // Process in batches
  for (let i = 0; i < contactIds.length; i += 100) {
    const batch = contactIds.slice(i, i + 100);
    batch.forEach(contactId => {
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
        
        updateContactDB(ss, {
          contactId: contactId,
          contactName: `${props.firstname || ''} ${props.lastname || ''}`.trim(),
          contactEmail: props.email || '',
          associatedCompany: props.company || '',
          phone: props.phone || '',
          jobTitle: props.jobtitle || ''
        });
      } catch (error) {
        Logger.log(`[REFRESH ERROR] Contact ${contactId}: ${error.message}`);
      }
    });
    Utilities.sleep(300);
  }
  
  Logger.log('[CONTACT DB] Refresh complete');
}

/**
 * Refresh all company data in Companies DB from HubSpot
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
  
  // Process in batches
  for (let i = 0; i < companyIds.length; i += 100) {
    const batch = companyIds.slice(i, i + 100);
    batch.forEach(companyId => {
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
        
        updateCompaniesDB(ss, {
          companyId: companyId,
          companyName: props.name || '',
          companyOwnerId: props.hubspot_owner_id || '',
          companySegment: props.segment || '',
          industry: props.industry || '',
          website: props.website || '',
          phone: props.phone || ''
        });
      } catch (error) {
        Logger.log(`[REFRESH ERROR] Company ${companyId}: ${error.message}`);
      }
    });
    Utilities.sleep(300);
  }
  
  Logger.log('[COMPANIES DB] Refresh complete');
}

/**
 * Refresh all owner data in Owner DB from HubSpot
 */
function refreshOwnerDB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Owner DB');
  
  if (!sheet) {
    Logger.log('[OWNER DB] No Owner DB sheet found');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      const ownerId = data[i][0].toString();
      fetchOwnerFromHubSpot(ss, ownerId);
      Utilities.sleep(100);
    }
  }
  
  Logger.log('[OWNER DB] Refresh complete');
}

// ============================================================================
// 8. INITIALIZATION FUNCTIONS
// ============================================================================

/**
 * Initialize enrichment system - creates DB sheets if they don't exist
 * If sheets are new, fetches all data from HubSpot
 */
function initEnrichmentDBs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  initializeEnrichmentDBs(ss);
  
  SpreadsheetApp.getUi().alert('Enrichment DBs initialized. Contact DB, Companies DB, and Owner DB sheets are ready.');
}
