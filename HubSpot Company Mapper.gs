/**
 * =============================================================================
 * HubSpot Data Mapper (Companies, Contacts, Owners & Emails) - PERFORMANCE OPTIMIZED
 * =============================================================================
 */
/**
 * Main Orchestrator updated with Cleanup Logic and Header Mapping
 */
function syncAllPipelineData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const progressSheet = ss.getSheetByName('Progress Sheet');
  
  try {
    const companyMapSheet = getOrCreateMapSheet(ss, 'Company_ID_Map', ['Company ID', 'Company Name']);
    const contactMapSheet = getOrCreateMapSheet(ss, 'Contact_ID_Map', ['Contact ID', 'Contact Name', 'Contact Email']);
    const ownerMapSheet   = getOrCreateMapSheet(ss, 'Owner_ID_Map', ['Owner ID', 'Owner Name']);

    let centralCompanyMap = getExistingMap(companyMapSheet, 2); 
    let centralContactMap = getExistingMap(contactMapSheet, 3); 
    let centralOwnerMap   = getExistingMap(ownerMapSheet, 2);   
    
    const pipelineNames = ['Payment Pipeline', 'Sales Pipeline', 'CS Pipeline'];
    
    pipelineNames.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;

      if (progressSheet) updateProgress(progressSheet, 'Status', `Syncing ${sheetName}...`);
      
      const fullRange = sheet.getDataRange();
      let data = fullRange.getValues();
      if (data.length < 3) return; 

      const headers = data[1]; 
      
      // Fixed Column Mappings based on your Provided Headers
      const COL_OWNER_ID = 5; // F: Owner ID
      const COL_COMP_ID  = 6; // G: Associated Company ID
      const COL_CONT_ID  = 7; // H: Associated Contact ID
      
      const nameCompIdx  = findOrCreateColumnIdx(sheet, headers, "Company Name", data);
      const nameContIdx  = findOrCreateColumnIdx(sheet, headers, "Contact Name", data);
      const emailContIdx = findOrCreateColumnIdx(sheet, headers, "Contact Email", data);
      const nameOwnerIdx = findOrCreateColumnIdx(sheet, headers, "Owner Name", data);

      const missingCompIds = [];
      const missingContIds = [];
      const missingOwnerIds = [];

      // Pass 1: Identification
      for (let i = 2; i < data.length; i++) {
        const compId  = data[i][COL_COMP_ID]?.toString().trim() || "";
        const contId  = data[i][COL_CONT_ID]?.toString().trim() || "";
        const ownerId = data[i][COL_OWNER_ID]?.toString().trim() || "";
        const currentCompName = data[i][nameCompIdx]?.toString() || "";

        // CLEANUP TRIGGER: If we have a Company ID but the name is currently a fallback (contains [Individual] or is an email)
        const isFallback = currentCompName.includes("[Individual]") || currentCompName.includes("@");
        const needsUpdate = !currentCompName || isFallback;

        // Logic for Contact
        if (contId && (!data[i][nameContIdx] || !data[i][emailContIdx])) {
          if (centralContactMap[contId]) {
            data[i][nameContIdx] = centralContactMap[contId].name;
            data[i][emailContIdx] = centralContactMap[contId].email;
          } else {
            missingContIds.push(contId);
          }
        }

        // Logic for Company (Includes Cleanup check)
        if (compId && compId !== "Individual Client" && needsUpdate) {
          if (centralCompanyMap[compId]) {
             data[i][nameCompIdx] = centralCompanyMap[compId];
          } else {
             missingCompIds.push(compId);
          }
        }
      }

      // API Fetching (Same as original)
      if (missingCompIds.length > 0) {
        const newNames = fetchBatchFromHubSpot(missingCompIds, 'companies', ['name'], (p) => p.name);
        Object.assign(centralCompanyMap, newNames);
        updateMapSheet(companyMapSheet, centralCompanyMap, 2);
      }
      // ... (Rest of missingContIds and missingOwnerIds fetch logic remains same)

      // Pass 3: Finalize with Hierarchy & Cleanup
      for (let i = 2; i < data.length; i++) {
        const compId = data[i][COL_COMP_ID]?.toString().trim();
        const currentCompName = data[i][nameCompIdx]?.toString() || "";

        // 1. Ensure Contact data is populated first for fallback use
        const contId = data[i][COL_CONT_ID]?.toString().trim();
        if (contId && !data[i][nameContIdx]) {
          data[i][nameContIdx] = centralContactMap[contId]?.name || "";
          data[i][emailContIdx] = centralContactMap[contId]?.email || "";
        }

        // 2. Finalize Company Name
        const isFallback = currentCompName.includes("[Individual]") || currentCompName.includes("@");
        
        if (!currentCompName || isFallback) {
          if (compId && compId !== "Individual Client" && centralCompanyMap[compId]) {
            // SUCCESS: Cleanup/Overwrite fallback with real Company Name
            data[i][nameCompIdx] = centralCompanyMap[compId];
          } else if (!compId || compId === "Individual Client") {
            // FALLBACK: Use Contact info
            const cName = data[i][nameContIdx];
            const cEmail = data[i][emailContIdx];

            if (cName && cName !== "N/A" && cName !== "") {
              data[i][nameCompIdx] = "[Individual](" + cName + ")";
            } else if (cEmail && cEmail !== "N/A" && cEmail !== "") {
              data[i][nameCompIdx] = cEmail;
            }
          }
        }
      }

      // BATCH WRITE
      sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    });
    // ... (Error handling remains same)
  } catch (e) { /* ... */ }
}
