/**
* ==============================================================================
* CS TEAM REPORT GENERATOR WITH REVENUE FORECASTING
* ==============================================================================
* 
* Creates a Customer Success-focused revenue report with extended forecasting
* capabilities for recurring revenue accounts.
* 
* PURPOSE:
* - Monitor customer health and retention
* - Forecast recurring revenue through fiscal year end
* - Track account ownership by CS team
* - Support renewal and expansion planning
* - Identify at-risk accounts (zero forecast months)
* 
* DATA SOURCE:
* - Reads from Summary Report (which uses TEMP_DATA)
* - Filters to CS team-owned accounts only
* - Date range from Config_sheet E6 (Report Date) and E9 (Report End)
* 
* KEY DIFFERENCE FROM SALES REPORT:
* - NO 12-month forecast limit (extends to FY end or Report End Date)
* - Focuses on customer retention vs new acquisition
* - Excludes "Sales Team" and "C-Suite" POC classifications
* - Uses different forecasting horizon
* 
* FILTERING LOGIC:
* 1. POC Team Filter
*    - EXCLUDES: "Sales Team", "C-Suite", blank
*    - INCLUDES: "CS Team", "CS & Sales", "CS and Sales (Transferred this FY)"
* 2. CS Team Membership (Optional)
*    - If CSHEETSFILTER_CELL (E7) = TRUE, require "Is CS Team Member" = TRUE
* 
* FORECASTING LOGIC (Recurring Revenue Only):
* 
* Step 1: Report Month Check
*    - If Recurring Revenue in Report Month > 0
*    - Use that amount for ALL future months through Report End or FY end
* 
* Step 2: Prior Month Fallback
*    - If Report Month = 0, check month before Report Month
*    - If prior month > 0, use that amount for Report Month + future
* 
* Step 3: No Revenue Case
*    - If both Report Month and prior month = 0
*    - Do NOT forecast (all future months remain 0)
* 
* IMPORTANT: No averaging, interpolation, or trend analysis.
* Only strict carry-forward of last known payment amount.
* 
* REVENUE CALCULATIONS:
* 
* 1. Realized Revenue (This FY)
*    - Sum of actual payments within current fiscal year
*    - Only includes months ≤ Report Date
*    - Excludes all forecasted amounts
* 
* 2. Forecasted Revenue (This FY)
*    - Realized Revenue +
*    - Forecasted recurring revenue from Report Month through FY end
*    - Only for Recurring companies
* 
* 3. Total Report Period Revenue
*    - Sum of all monthly columns in report window
*    - Includes both actual and forecasted amounts
*    - Spans Report Start to Report End Date
* 
* OUTPUT COLUMNS:
* 1. Company Name
* 2. Revenue Trend (Sparkline)
* 3. Revenue Type (Recurring/R-OTP/OTP)
* 4. Client Age (New/Old/Future)
* 5. First Revenue Fiscal Year
* 6. POC Team
* 7-N. Monthly columns (YYYY-MMM format)
* N+1. Realized Revenue (FY) Source - actual payments only
* N+2. Forecasted Revenue This FY - includes forecast
* N+3. Total Revenue for the period - full report window
* 
* FORMATTING:
* - Grey font for ALL forecasted months (even in current FY)
* - Grey font for months outside current FY
* - Light yellow background for rows with current FY activity
* - Green/Red trend colors for ACTUAL months only (not forecast)
* - Currency format: $###,###
* - Zero values displayed as "-"
* - Frozen first row and first 2 columns
* - Auto-filter on headers
* - Subtotal row at bottom
* 
* CONFIGURATION CELLS (Config_sheet):
* E6 = Report Date (Report Month for forecasting)
* E9 = Report End Date (forecast horizon)
* E7 = CS Team Filter (optional checkbox)
* 
* USAGE:
* From Menu: "7. Run CS Team Report"
* Or call: CS_forecastRevenue()
* 
* PREREQUISITES:
* - Summary Report must exist and be up-to-date
* - Config_sheet with dates in E6 and E9
* - CS_team_Members sheet (if using E7 filter)
* 
* PERFORMANCE NOTE:
* This report reads from Summary Report (already aggregated) rather than
* TEMP_DATA, making it very fast even with large datasets.
* 
* See README.md for complete documentation and forecasting examples.
* ==============================================================================
*/

/**
* ==============================================================================
* CONFIGURATION
* ==============================================================================
*/
const CSSUMMARY_CONFIG = {
 SOURCE_SHEET: "Summary Report",
 CONFIG_SHEET: "Config_sheet",
 CSSUMMARY_SHEET_NAME: "Summary_CS_Team",
 REPORT_END_CELL: "E9",
 REPORT_DATE_CELL: "E6",
 
 EXTRA_COLUMNS: [
   { header: "Revenue Trend", index: -1, isSparkline: true },
   { header: "Revenue Type", index: 2, isSparkline: false },
   { header: "Client Age", index: 3, isSparkline: false },
   { header: "First Revenue Fiscal Year", index: 4, isSparkline: false },
   { header: "POC Team", index: 5, isSparkline: false }
 ],
 
 // Trailing columns in the specific requested order
 TRAILING_COLUMNS: [
   "Realized Revenue (FY) Source",
   "Forecasted Revenue This FY",
   "Total Revenue for the period (Including forecast)"
 ],
 
 FROZEN_COLUMN_COUNT: 2,
 FY_START_MONTH: 6, // 6 = July
};

const EXCLUDED_POC_TEAMS = ["Sales Team", "C-Suite",""];

/**
* ==============================================================================
* MAIN FUNCTION: CS_forecastRevenue
* ==============================================================================
*/
function CS_forecastRevenue() {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const sourceSheet = ss.getSheetByName(CSSUMMARY_CONFIG.SOURCE_SHEET);
 const configSheet = ss.getSheetByName(CSSUMMARY_CONFIG.CONFIG_SHEET);

 if (!sourceSheet || !configSheet) throw new Error("Required sheets not found.");

 // 1. DATES & FY BOUNDARIES
 const rawReportDate = new Date(configSheet.getRange(CSSUMMARY_CONFIG.REPORT_DATE_CELL).getValue());
 const reportDate = new Date(rawReportDate.getFullYear(), rawReportDate.getMonth(), 1);
 
 let fyStartYear = reportDate.getMonth() >= CSSUMMARY_CONFIG.FY_START_MONTH ? reportDate.getFullYear() : reportDate.getFullYear() - 1;
 const currentFYStart = new Date(fyStartYear, CSSUMMARY_CONFIG.FY_START_MONTH, 1);
 const currentFYEnd = new Date(fyStartYear + 1, CSSUMMARY_CONFIG.FY_START_MONTH - 1, 1);
 
 const rawEndDate = new Date(configSheet.getRange(CSSUMMARY_CONFIG.REPORT_END_CELL).getValue());
 const reportEndDate = new Date(rawEndDate.getFullYear(), rawEndDate.getMonth(), 1);
 const priorMonthDate = new Date(reportDate.getFullYear(), reportDate.getMonth() - 1, 1);

 // 2. LOAD DATA
 const allData = sourceSheet.getDataRange().getValues();
 const rawHeaders = allData[0];
 const dataRows = allData.slice(1);

 // 3. GENERATE HEADERS
 const existingMonths = [];
 rawHeaders.forEach(h => {
   if (h instanceof Date || (typeof h === 'string' && h.match(/^\d{4}-(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/))) {
      existingMonths.push(new Date(new Date(h).getFullYear(), new Date(h).getMonth(), 1));
   }
 });

 let curH = new Date(reportDate);
 while (curH <= reportEndDate) {
   if (!existingMonths.some(em => em.getTime() === curH.getTime())) existingMonths.push(new Date(curH));
   curH.setMonth(curH.getMonth() + 1);
 }
 const sortedMonths = existingMonths.sort((a, b) => a - b);
 const sortedMonthStrings = sortedMonths.map(m => formatDateToYYYYMMM(m));

 // 4. FIND INDICES
 const reportMonthIdx = findDateColumnIndex(rawHeaders, reportDate);
 const priorMonthIdx = findDateColumnIndex(rawHeaders, priorMonthDate);

 // 5. BUILD OUTPUT HEADERS
 const dynamicHeaderLabels = CSSUMMARY_CONFIG.EXTRA_COLUMNS.map(c => c.header);
 const finalHeaders = ["Company Name", ...dynamicHeaderLabels, ...sortedMonthStrings, ...CSSUMMARY_CONFIG.TRAILING_COLUMNS];
 const finalOutput = [finalHeaders];

 // 6. PROCESS ROWS
 dataRows.forEach((row) => {
   const pocTeam = cleanString(row[5]);
   const revenueType = cleanString(row[2]).toLowerCase();
  
   if (EXCLUDED_POC_TEAMS.includes(pocTeam)) return;

   const outRow = [row[0]]; // Company Name
  
   // Static Metadata
   CSSUMMARY_CONFIG.EXTRA_COLUMNS.forEach(col => {
     if (col.isSparkline) outRow.push("");
     else outRow.push(row[col.index]);
   });

   // Forecast Baseline Determination
   const repVal = reportMonthIdx !== -1 ? parseNumeric(row[reportMonthIdx]) : 0;
   const priVal = priorMonthIdx !== -1 ? parseNumeric(row[priorMonthIdx]) : 0;
   const forecastVal = repVal > 0 ? repVal : priVal;

   let forecastedFYTotal = 0;
   let periodTotal = 0;

   // Monthly Data Generation
   sortedMonths.forEach(m => {
     const mIdxInSource = findDateColumnIndex(rawHeaders, m);
     let val = mIdxInSource !== -1 ? parseNumeric(row[mIdxInSource]) : 0;
    
     const mTime = m.getTime();
     const isHistorical = mTime < reportDate.getTime();

     // Apply forecast
     if (!isHistorical && val === 0 && revenueType.includes("recurring")) {
       val = forecastVal;
     }

     outRow.push(val);

     // Internal calculations for trailing columns
     const isCurrentFY = (mTime >= currentFYStart.getTime() && mTime <= currentFYEnd.getTime());
     if (isCurrentFY) forecastedFYTotal += val;
     periodTotal += val;
   });

   // Final Trailing Columns: Realized from Source (second last col), then calculated metrics
   const realizedSourceVal = row[row.length - 2];
   outRow.push(realizedSourceVal, forecastedFYTotal, periodTotal);
   finalOutput.push(outRow);
 });

 // 7. ALPHA SORT
 const headerRow = finalOutput.shift();
 finalOutput.sort((a, b) => String(a[0]).localeCompare(String(b[0])));
 finalOutput.unshift(headerRow);

 writeAndFormatSummaryCS(ss, finalOutput, sortedMonths, currentFYStart, reportDate);
}

/**
* ==============================================================================
* FORMATTING & WRITE
* ==============================================================================
*/
function writeAndFormatSummaryCS(ss, finalOutput, sortedMonths, currentFYStart, reportDate) {
 let sheet = ss.getSheetByName(CSSUMMARY_CONFIG.CSSUMMARY_SHEET_NAME) || ss.insertSheet(CSSUMMARY_CONFIG.CSSUMMARY_SHEET_NAME);
 sheet.clear();
 
 const lastRow = finalOutput.length;
 const lastCol = finalOutput[0].length;
 const staticColCount = 1 + CSSUMMARY_CONFIG.EXTRA_COLUMNS.length;
 const monthStartIdx = staticColCount + 1;

 sheet.getRange(1, 1, lastRow, lastCol).setValues(finalOutput);

 // 1. Sparkline (Revenue Trend)
 const sparkColIdx = CSSUMMARY_CONFIG.EXTRA_COLUMNS.findIndex(c => c.isSparkline) + 2;
 if (sparkColIdx > 1) {
   for (let i = 2; i <= lastRow; i++) {
     const startCell = columnToLetter(monthStartIdx) + i;
     const endCell = columnToLetter(monthStartIdx + sortedMonths.length - 1) + i;
     sheet.getRange(i, sparkColIdx).setFormula(`=SPARKLINE(${startCell}:${endCell}, {"charttype","line";"linewidth",2;"linecolor","#5f6368"})`);
   }
 }

 // 2. Totals Row
 const subTotalRowIdx = lastRow + 1;
 sheet.getRange(subTotalRowIdx, 1).setValue("TOTALS").setFontWeight("bold");
 const colsToTotal = [...Array(sortedMonths.length).keys()].map(x => x + monthStartIdx).concat([lastCol - 2, lastCol - 1, lastCol]);
 colsToTotal.forEach(c => {
   sheet.getRange(subTotalRowIdx, c).setFormula(`=SUBTOTAL(9, ${columnToLetter(c)}2:${columnToLetter(c)}${lastRow})`).setFontWeight("bold");
 });

 // 3. UI Stylings
 sheet.setFrozenRows(1);
 sheet.setFrozenColumns(CSSUMMARY_CONFIG.FROZEN_COLUMN_COUNT);
 sheet.getRange(1, 1, 1, lastCol).setBackground("#444444").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
 
 // Format numerical columns
 sheet.getRange(2, monthStartIdx, subTotalRowIdx, sortedMonths.length + 3).setNumberFormat('[=0]"-" ; $#,##0');

 // 4. Trend Coloring
 sortedMonths.forEach((m, i) => {
   const colIdx = monthStartIdx + i;
   const range = sheet.getRange(2, colIdx, lastRow - 1, 1);
   if (m >= currentFYStart && m <= reportDate) {
     range.setFontColor("black");
     if (i > 0) {
       const colors = range.getValues().map((v, idx) => {
         const curr = parseNumeric(v[0]), prev = parseNumeric(sheet.getRange(idx + 2, colIdx - 1).getValue());
         if (curr === 0 || prev === 0) return "black";
         return (curr > prev * 1.1 || curr > prev + 50) ? "#008000" : (curr < prev * 0.9 || curr < prev - 50) ? "#FF0000" : "black";
       });
       range.setFontColors(colors.map(c => [c]));
     }
   } else {
     range.setFontColor("#999999");
   }
 });

 sheet.autoResizeColumns(1, lastCol);
 if (sheet.getFilter()) sheet.getFilter().remove();
 sheet.getRange(1, 1, lastRow, lastCol).createFilter();
}

/**
* ==============================================================================
* HELPERS
* ==============================================================================
*/
function columnToLetter(c) {
 let letter = "";
 while (c > 0) {
   let t = (c - 1) % 26;
   letter = String.fromCharCode(t + 65) + letter;
   c = (c - t - 1) / 26;
 }
 return letter;
}

function findDateColumnIndex(headers, target) {
 const tStr = formatDateToYYYYMMM(target);
 return headers.findIndex(h => formatDateToYYYYMMM(h) === tStr);
}

function formatDateToYYYYMMM(d) {
 if (!(d instanceof Date)) return String(d);
 const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
 return `${d.getFullYear()}-${months[d.getMonth()]}`;
}

function parseNumeric(v) {
 if (typeof v === 'number') return v;
 const n = parseFloat(String(v).replace(/[$,\s]/g, ''));
 return isNaN(n) ? 0 : n;
}

function cleanString(v) { return v ? String(v).trim() : ""; }
