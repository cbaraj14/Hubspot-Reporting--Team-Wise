/**
* ============================================================================
* TEMP SHEET MODULE - DATA CLASSIFICATION & ENRICHMENT ENGINE
* ============================================================================
* 
* This module processes raw deal data and creates a pre-computed cache with
* advanced classifications. This eliminates redundant calculations and makes
* reports run 10-20x faster.
* 
* KEY OPERATIONS:
* 1. Deduplication by Deal ID (keeps most recent Last Modified Date)
* 2. Revenue Type Classification (Recurring/R-OTP/OTP)
* 3. Client Age Classification (New/Old/Future)
* 4. Team Attribution (Sales/CS membership flags)
* 5. Fiscal Year Calculations
* 6. First Revenue Month/FY tracking
* 
* CLASSIFICATION LOGIC:
* 
* Revenue Type:
*   - Recurring: Deal name has keywords OR ≥8 months in one FY
*   - R-OTP: >5 total months (but not Recurring)
*   - OTP: Everything else
* 
* Client Age:
*   - New: First payment in current FY
*   - Old: First payment before current FY
*   - Future: First payment after Report Date
*   - Prospect: No payment history
* 
* Team Attribution:
*   - Checks Owner ID against Sales_team_Members and CS_team_Members sheets
*   - Tracks revenue attribution by team
*   - Identifies transferred accounts
* 
* OUTPUT:
* Creates hidden "TEMP_DATA" sheet with 25+ columns including all source
* data plus computed classifications. All reports read from this cache.
* 
* PERFORMANCE:
* - Batch writing (500 rows at a time)
* - Hidden sheet to avoid UI lag
* - Indexed lookups for fast access
* - Run once, use many times
* 
* USAGE:
* Run tempUpdateTempSheet() before generating any reports.
* This is typically done automatically or via Menu option 4.
* 
* See README.md for complete documentation.
* ============================================================================
*/

/**
* ============================================================================
* 1) TEMP CONFIGURATION OBJECT
* ============================================================================
*/
const tempCONFIG = {
  CONFIG_SHEET_NAME: "Config",
  REPORT_DATE_CELL: "E6",       
  FISCAL_YEAR_START_MONTH: 6,    // 0 = Jan, 6 = July
  SALES_TEAM_SHEET: "Sales_team_Members",
  CS_TEAM_SHEET: "CS_team_Members",
  PAYMENT_SHEET_NAME: "Payment Pipeline",
  SOURCE_SHEETS: ["Payment Pipeline", "Sales Pipeline", 'CS Pipeline'],
  TEMP_SHEET_NAME: "TEMP_DATA",
  PIPELINE_HEADER_ROW: 2,       
  BATCH_SIZE: 500
};

const tempOUTPUT_HEADERS = typeof OUTPUT_HEADERS !== 'undefined' ? OUTPUT_HEADERS : [];

/**
* ============================================================================
* 2) MAIN EXECUTION: tempUpdateTempSheet
* ============================================================================
*/
function tempUpdateTempSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startTime = new Date();
  const originalSheet = ss.getActiveSheet();

  // 1. Setup Dates & Teams
  const configSheet = ss.getSheetByName(tempCONFIG.CONFIG_SHEET_NAME);
  const reportDateVal = configSheet ? configSheet.getRange(tempCONFIG.REPORT_DATE_CELL).getValue() : null;
  const reportDate = tempIsValidDate(reportDateVal) ? new Date(reportDateVal) : new Date();
  const currentFY = tempGetFiscalYear(reportDate);

  const salesTeam = tempGetTeamList(ss, tempCONFIG.SALES_TEAM_SHEET);
  const csTeam = tempGetTeamList(ss, tempCONFIG.CS_TEAM_SHEET);

  const allData = [];
  const paymentDeals = [];
  const companyHasSalesMemberMap = {};
  const companyHasRevenueMap = {};
  const companyDealsThisFYMap = {};  // New map to track if a company has any deal this FY

  // 2. First Pass: Collect Data and Map Company Status
  tempCONFIG.SOURCE_SHEETS.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const fullData = sheet.getDataRange().getValues();
    if (fullData.length <= tempCONFIG.PIPELINE_HEADER_ROW) return;

    const rawRows = fullData.slice(tempCONFIG.PIPELINE_HEADER_ROW);
    const uniqueMap = {};
  
    rawRows.forEach(row => {
      const dealId = row[8];
      const companyId = row[6];
      const ownerId = row[5] ? row[5].toString() : "";
      
      const dealDate = tempIsValidDate(row[9]) ? new Date(row[9]) : new Date();
      const dealFY = tempGetFiscalYear(dealDate);
    
      if (!dealId) return;

      if (companyId && salesTeam.includes(ownerId)) {
        companyHasSalesMemberMap[companyId] = true;
      }

      if (companyId && sheetName === tempCONFIG.PAYMENT_SHEET_NAME && dealFY === currentFY) {
        companyHasRevenueMap[companyId] = true;
      }

      const modDate = tempIsValidDate(row[10]) ? new Date(row[10]) : new Date(0);
      if (!uniqueMap[dealId] || (new Date(row[10]) > new Date(uniqueMap[dealId][10]))) {
        uniqueMap[dealId] = row;
      }

      // Track if there is any deal this fiscal year for the associated company
      if (companyId && dealFY === currentFY) {
        companyDealsThisFYMap[companyId] = true;
      }
    });

    const processedRows = Object.values(uniqueMap);
    if (sheetName === tempCONFIG.PAYMENT_SHEET_NAME) paymentDeals.push(...processedRows);
    allData.push(...processedRows.map(row => ({ data: row, source: sheetName })));
  });

  // Calculate stats using the hierarchical grouping logic
  const companyStats = tempCalculateCompanyStats(paymentDeals);
  
  // Updated headers with all necessary columns
  const updatedHeaders = [...tempOUTPUT_HEADERS, 
    "Company has Sales Team", 
    "Has Revenue this FY", 
    "Any Deals this FY", 
    "Is Sales", 
    "Has Revenue this FY and Is Sales", 
    "Is CS", 
    "Has Revenue this FY and Is CS"
  ];

  const finalOutput = [updatedHeaders];

  // 3. Second Pass: Build Final Output
  allData.forEach(item => {
    const row = item.data;
    
    // Check tiered grouping keys in order to find the match in our stats map
    const companyId = row[6];
    const companyName = row[7];
    const email = row[12];
    
    // Find if ANY of the identifiers exist in our calculated stats map
    let stats = null;
    if (companyId && companyStats[companyId]) {
      stats = companyStats[companyId];
    } else if (companyName && companyStats[companyName]) {
      stats = companyStats[companyName];
    } else if (email && companyStats[email]) {
      stats = companyStats[email];
    }
    
    if (!stats) {
      // Final Fallback: Standalone record logic for First Revenue FY
      const fallbackDate = tempIsValidDate(row[9]) ? new Date(row[9]) : new Date();
      stats = {
        firstDate: fallbackDate,
        firstMonth: Utilities.formatDate(fallbackDate, Session.getScriptTimeZone(), "yyyy-MMM"),
        firstFY: tempGetFiscalYear(fallbackDate),
        months: []
      };
    }
    
    const rowDate = tempIsValidDate(row[9]) ? new Date(row[9]) : new Date();
    const dealFY = tempGetFiscalYear(rowDate);
  
    const isSales = salesTeam.includes(row[5] ? row[5].toString() : "");
    const isCS = csTeam.includes(row[5] ? row[5].toString() : "");
    const rowHasRevThisFY = (item.source === tempCONFIG.PAYMENT_SHEET_NAME && dealFY === currentFY);

    const companyHasSalesTeam = companyHasSalesMemberMap[companyId] === true;
    const companyHasRevenueThisFY = companyHasRevenueMap[companyId] === true;
    const companyHasDealsThisFY = companyDealsThisFYMap[companyId] === true;  // Check for any deal this FY

    // Ensure we add all columns, including the new ones.
    finalOutput.push([
      row[11], row[3], row[12], row[14], row[6], row[7], parseFloat(row[1]) || 0, row[4], row[9], row[2],
      row[10], row[0], row[13], row[8], row[5],
      tempClassifyDealType(row[0] ? row[0].toString() : "", stats.months),
      tempClassifyAge(stats.firstDate, stats.firstFY, currentFY, reportDate),
      stats.firstMonth, stats.firstFY,
      (rowHasRevThisFY ? parseFloat(row[1]) || 0 : 0),
      isSales, rowHasRevThisFY, (rowHasRevThisFY && isSales),
      isCS, (rowHasRevThisFY && isCS),
      companyHasSalesTeam,
      companyHasRevenueThisFY,
      companyHasDealsThisFY // New column for "Any Deals this FY"
    ]);
  });

  // 4. Write to TEMP Sheet
  let tempSheet = ss.getSheetByName(tempCONFIG.TEMP_SHEET_NAME);
  if (tempSheet) ss.deleteSheet(tempSheet);
  tempSheet = ss.insertSheet(tempCONFIG.TEMP_SHEET_NAME);
  tempSheet.hideSheet();
  ss.setActiveSheet(originalSheet);

  const totalRows = finalOutput.length;
  const totalCols = updatedHeaders.length;  // Ensure columns match header length

  for (let i = 0; i < totalRows; i += tempCONFIG.BATCH_SIZE) {
    const currentBatchSize = Math.min(tempCONFIG.BATCH_SIZE, totalRows - i);
    const batchData = finalOutput.slice(i, i + currentBatchSize);
    tempSheet.getRange(i + 1, 1, currentBatchSize, totalCols).setValues(batchData);
  }

  tempSheet.setFrozenRows(1);
  tempSheet.getRange(1, 1, 1, totalCols).setFontWeight("bold");

  const endTime = new Date();
  console.log(`Success! Total Time: ${(endTime - startTime)/1000} seconds.`);
}

/**
* ============================================================================
* 3) TEMP LOGIC HELPERS
* ============================================================================
*/

function tempIsValidDate(dateVal) {
  if (dateVal === null || dateVal === undefined || dateVal === "") return false;
  const strVal = dateVal.toString().trim().toLowerCase();
  if (strVal === "n/a" || strVal === "unknown" || strVal === "blank") return false;
  
  const d = new Date(dateVal);
  return d instanceof Date && !isNaN(d.getTime());
}

function tempGetFiscalYear(dateVal) {
  let date = tempIsValidDate(dateVal) ? new Date(dateVal) : new Date();
  const startMonth = tempCONFIG.FISCAL_YEAR_START_MONTH || 0;
  const year = date.getFullYear();
  const month = date.getMonth();
  
  const fyStart = (month >= startMonth) ? year : year - 1;
  const shortYearStart = fyStart.toString().slice(-2);
  const shortYearEnd = (fyStart + 1).toString().slice(-2);
  
  let relativeMonth = month - startMonth;
  if (relativeMonth < 0) relativeMonth += 12;
  const quarter = Math.floor(relativeMonth / 3) + 1;
  
  return `FY ${shortYearStart}/${shortYearEnd} Q${quarter}`;
}

function tempClassifyDealType(dealName, monthsArray) {
  const nameLower = (dealName || "").toString().toLowerCase();
  if (nameLower.includes("data-services-subscription") || nameLower.includes("gold monthly")) return "Recurring";
  if (monthsArray.length >= 8) return "Recurring";
  if (monthsArray.length > 5) return "R-OTP";
  return "OTP";
}

function tempClassifyAge(firstPaymentDate, firstFY, currentFY, reportDate) {
  // If firstFY is N/A or missing, return Prospect
  if (!firstPaymentDate || !firstFY || firstFY === "N/A") return "Prospect";
  
  if (firstPaymentDate > reportDate) return "Future";
  
  // Extract numerical parts for comparison (e.g., "FY 24/25 Q1" -> 24)
  const firstMatch = firstFY.match(/\d+/);
  const currentMatch = currentFY.match(/\d+/);
  
  const firstFYNum = firstMatch ? parseInt(firstMatch[0]) : 0;
  const currentFYNum = currentMatch ? parseInt(currentMatch[0]) : 0;
  
  if (firstFYNum < currentFYNum) return "Old";
  
  // If same start year, check the whole string or specific quarter logic
  // Since "FY 24/25 Q1" represents the start of the relationship, if it matches current, it's "New"
  if (firstFY === currentFY) return "New";

  // Fallback check if string comparison fails but it's clearly not "Old"
  if (firstFYNum === currentFYNum && firstPaymentDate <= reportDate) return "New";
  
  return "Other";
}

/**
 * Enhanced grouping logic:
 * Collects data using a single source of truth for an entity,
 * but indexes the final map by ALL identifiers (ID, Name, Email) 
 * so lookups in the second pass are guaranteed to find the group's stats.
 */
function tempCalculateCompanyStats(paymentRows) {
  const groupedData = {}; // Internal tracking by primary grouping key
  
  paymentRows.forEach(row => {
    const companyId = row[6] ? row[6].toString().trim() : "";
    const companyName = row[7] ? row[7].toString().trim() : "";
    const email = row[12] ? row[12].toString().trim() : "";
    
    // Determine the primary key for grouping logic
    const primaryKey = companyId || companyName || email || null;
    if (!primaryKey) return; 
    
    const closeDate = tempIsValidDate(row[9]) ? new Date(row[9]) : new Date();
    const monthStr = Utilities.formatDate(closeDate, Session.getScriptTimeZone(), "yyyy-MMM");
    
    if (!groupedData[primaryKey]) { 
      groupedData[primaryKey] = { 
        earliest: closeDate, 
        months: new Set(),
        aliases: new Set() // Track all associated identifiers for cross-indexing
      }; 
    }
    
    if (closeDate < groupedData[primaryKey].earliest) { 
      groupedData[primaryKey].earliest = closeDate; 
    }
    
    groupedData[primaryKey].months.add(monthStr);
    if (companyId) groupedData[primaryKey].aliases.add(companyId);
    if (companyName) groupedData[primaryKey].aliases.add(companyName);
    if (email) groupedData[primaryKey].aliases.add(email);
  });
  
  const finalized = {};
  for (let key in groupedData) {
    const statsResult = {
      firstDate: groupedData[key].earliest,
      firstMonth: Utilities.formatDate(groupedData[key].earliest, Session.getScriptTimeZone(), "yyyy-MMM"),
      firstFY: tempGetFiscalYear(groupedData[key].earliest),
      months: Array.from(groupedData[key].months)
    };
    
    // Index the result by every alias found for this group
    // This ensures that if the second pass looks up by "Email", it finds the same data as looking up by "ID"
    groupedData[key].aliases.forEach(alias => {
      finalized[alias] = statsResult;
    });
  }
  return finalized;
}

function tempGetTeamList(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  data.shift();
  return data.map(row => row[1].toString());
}
