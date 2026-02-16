/**
 * ============================================================================
 * 1) CONFIGURATION (MODULAR CONTROL PANEL)
 * ============================================================================
 */
const SUMMARY_CONFIG = {
  // UPDATED: Now points to your "temp sheet"
  SOURCE_SHEET: "TEMP_DATA", 
  CONFIG_SHEET: "Config_sheet",
  
  REPORT_START_CELL: "E8",
  REPORT_END_CELL: "E9",
  
  Summary_Sheet: "Summary Report", 
  TARGET_PIPELINE: "Payment Pipeline",
  FY_START_MONTH: 6, 
  
  FROZEN_COLUMN_COUNT: 3,

  EXTRA_COLUMNS: [
    { header: "Revenue Trend", isSparkline: true },
    { header: "Revenue Type", index: 15 },
    { header: "Client Age", index: 16 },
    { header: "First Revenue Fiscal Year", index: 18 },
    { header: "POC Team", isVirtual: true },
  ],

  COL_COMPANY: 0,
  COL_AMOUNT: 6,
  COL_PIPELINE: 7,
  COL_CLOSE_DATE: 8,
  COL_IS_SALES: 20, 
  COL_IS_CS: 23
};

/**
 * ============================================================================
 * 2) MAIN EXECUTION
 * ============================================================================
 */
function generateSummaryReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Attempt to grab the sheets
  const sourceSheet = ss.getSheetByName(SUMMARY_CONFIG.SOURCE_SHEET);
  const configSheet = ss.getSheetByName(SUMMARY_CONFIG.CONFIG_SHEET);
  
  // CRITICAL CHECK: If the input data is missing, we must stop.
  if (!sourceSheet) {
    throw new Error(`CRITICAL ERROR: The source data sheet "${SUMMARY_CONFIG.SOURCE_SHEET}" was not found. Please ensure it exists.`);
  }
  if (!configSheet) {
    throw new Error(`CRITICAL ERROR: The configuration sheet "${SUMMARY_CONFIG.CONFIG_SHEET}" was not found.`);
  }

  const startVal = configSheet.getRange(SUMMARY_CONFIG.REPORT_START_CELL).getValue();
  const endVal = configSheet.getRange(SUMMARY_CONFIG.REPORT_END_CELL).getValue();
  
  if (!startVal || !endVal) {
    SpreadsheetApp.getUi().alert("Error: Dates are missing in E8/E9 of the Config sheet.");
    return;
  }

  const startDate = new Date(startVal);
  const endDate = new Date(endVal);
  const reportFY = getFYString(endDate);

  const data = sourceSheet.getDataRange().getValues();
  const rows = data.slice(1);
  
  const pivotMap = {};
  const allMonths = new Set();
  const globalHistory = {}; 

  // A) DATA AGGREGATION
  rows.forEach(row => {
    const company = row[SUMMARY_CONFIG.COL_COMPANY];
    if (!company) return;

    if (!globalHistory[company]) {
      globalHistory[company] = { allSales: true, allCS: true, anySales: false, anyCS: false };
    }
    if (row[SUMMARY_CONFIG.COL_IS_SALES]) globalHistory[company].anySales = true; else globalHistory[company].allSales = false;
    if (row[SUMMARY_CONFIG.COL_IS_CS]) globalHistory[company].anyCS = true; else globalHistory[company].allCS = false;

    const closeDate = new Date(row[SUMMARY_CONFIG.COL_CLOSE_DATE]);
    if (row[SUMMARY_CONFIG.COL_PIPELINE] === SUMMARY_CONFIG.TARGET_PIPELINE && closeDate >= startDate && closeDate <= endDate) {
      const monthKey = Utilities.formatDate(closeDate, Session.getScriptTimeZone(), "yyyy-MMM");
      allMonths.add(monthKey);

      if (!pivotMap[company]) {
        pivotMap[company] = {
          name: company, tempRow: row, monthlyValues: {}, realizedFY: 0, totalRevenue: 0,
          payAllSales: true, payAllCS: true, fyHasSales: false, fyHasCS: false, hasFYRev: false
        };
      }

      const p = pivotMap[company];
      const amount = parseFloat(row[SUMMARY_CONFIG.COL_AMOUNT]) || 0;
      p.monthlyValues[monthKey] = (p.monthlyValues[monthKey] || 0) + amount;
      p.totalRevenue += amount;
      if (!row[SUMMARY_CONFIG.COL_IS_SALES]) p.payAllSales = false;
      if (!row[SUMMARY_CONFIG.COL_IS_CS]) p.payAllCS = false;

      if (getFYString(closeDate) === reportFY) {
        p.realizedFY += amount;
        p.hasFYRev = true;
        if (row[SUMMARY_CONFIG.COL_IS_SALES]) p.fyHasSales = true;
        if (row[SUMMARY_CONFIG.COL_IS_CS]) p.fyHasCS = true;
      }
    }
  });

  // B) HEADERS & SORTING
  const sortedMonths = Array.from(allMonths).sort((a,b) => new Date(a) - new Date(b));
  const dynamicHeaders = SUMMARY_CONFIG.EXTRA_COLUMNS.map(c => c.header);
  const headers = ["Company Name", ...dynamicHeaders, ...sortedMonths, "Realized Revenue (FY)", "Total Revenue"];
  
  let sortedData = Object.values(pivotMap).sort((a, b) => {
    const fyA = a.tempRow[18] || "";
    const fyB = b.tempRow[18] || "";
    return fyA !== fyB ? fyA.localeCompare(fyB) : a.name.localeCompare(b.name);
  });

  // C) BUILD ARRAY
  const finalOutput = [headers];
  sortedData.forEach(p => {
    const row = [p.name];
    SUMMARY_CONFIG.EXTRA_COLUMNS.forEach(col => {
      if (col.isSparkline) row.push(""); 
      else if (col.isVirtual) row.push(calculatePOCTag(p, globalHistory[p.name]));
      else row.push(p.tempRow[col.index]);
    });
    sortedMonths.forEach(m => row.push(p.monthlyValues[m] || 0));
    row.push(p.realizedFY, p.totalRevenue);
    finalOutput.push(row);
  });

  // D) WRITE & FORMAT
  writeAndFormatSummary(ss, finalOutput, sortedMonths, reportFY);
}

/**
 * ============================================================================
 * 3) SUPPORTING LOGIC
 * ============================================================================
 */
function writeAndFormatSummary(ss, finalOutput, sortedMonths, reportFY) {
  //let sheet = ss.getSheetByName(SUMMARY_CONFIG.Summary_Sheet);
  let sheet = '';
  sheet = ss.getSheetByName(SUMMARY_CONFIG.Summary_Sheet);
  console.log(sheet)
  // If the Summary Report was deleted, this creates it again
  if (!sheet) {
    sheet = ss.insertSheet(SUMMARY_CONFIG.Summary_Sheet);
  }
  
  // Wipes content and formatting to start fresh
  sheet.clear();
  
  const lastRow = finalOutput.length;
  if (lastRow === 1) {
    SpreadsheetApp.getUi().alert("No data found for the selected date range.");
    return;
  }
  
  const lastCol = finalOutput[0].length;
  const staticColCount = 1 + SUMMARY_CONFIG.EXTRA_COLUMNS.length;
  const monthStartIdx = staticColCount + 1;
  const monthEndIdx = staticColCount + sortedMonths.length;

  sheet.getRange(1, 1, lastRow, lastCol).setValues(finalOutput);

  // 1. INJECT FORMULAS
  const sparklineConfigIdx = SUMMARY_CONFIG.EXTRA_COLUMNS.findIndex(c => c.isSparkline);
  if (sparklineConfigIdx !== -1) {
    const sparkColIdx = sparklineConfigIdx + 2; 
    for (let i = 2; i <= lastRow; i++) {
      const startRange = columnToLetter(monthStartIdx) + i;
      const endRange = columnToLetter(monthEndIdx) + i;
      sheet.getRange(i, sparkColIdx).setFormula(`=SPARKLINE(${startRange}:${endRange}, {"charttype","line";"linewidth",2;"linecolor","#5f6368"})`);
    }
  }

  const subTotalIdx = lastRow + 1;
  sheet.getRange(subTotalIdx, 1, 1, staticColCount).setValues([["TOTALS", ...new Array(staticColCount-1).fill("")]]);
  
  [...Array(sortedMonths.length).keys()].map(x => x + monthStartIdx).concat([lastCol - 1, lastCol]).forEach(c => {
    const colLetter = columnToLetter(c);
    sheet.getRange(subTotalIdx, c).setFormula(`=SUBTOTAL(9, ${colLetter}2:${colLetter}${lastRow})`).setFontWeight("bold");
  });

  // 2. FORMATTING
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(SUMMARY_CONFIG.FROZEN_COLUMN_COUNT);
  sheet.getRange(1, 1, 1, lastCol).setBackground("#444444").setFontColor("white").setFontWeight("bold");
  
  if (sortedMonths.length > 0) {
    sheet.getRange(2, monthStartIdx, subTotalIdx, sortedMonths.length).setNumberFormat('[=0]"-" ; $#,##0.00');
  }
  sheet.getRange(2, lastCol - 1, subTotalIdx, 2).setNumberFormat('$#,##0.00');

  // Trend Coloring
  sortedMonths.forEach((m, i) => {
    const colIdx = monthStartIdx + i;
    const range = sheet.getRange(2, colIdx, lastRow - 1, 1);
    if (getFYString(new Date(m)) === reportFY && i > 0) {
      const vals = range.getValues(), prevVals = sheet.getRange(2, colIdx - 1, lastRow - 1, 1).getValues();
      const colors = vals.map((v, idx) => {
        const curr = v[0], prev = prevVals[idx][0];
        if (curr === 0 || prev === 0) return "black";
        return (curr > prev * 1.1 || curr > prev + 50) ? "green" : (curr < prev * 0.9 || curr < prev - 50) ? "red" : "black";
      });
      range.setFontColors(colors.map(c => [c]));
    } else if (getFYString(new Date(m)) !== reportFY) range.setFontColor("#999999");
  });

  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, lastRow, lastCol).createFilter();
  sheet.autoResizeColumns(1, lastCol);
}

function calculatePOCTag(p, g) {
  if (p.payAllSales && !p.payAllCS) return "Sales Team";
  if (p.payAllCS && !p.payAllSales) return "CS Team";
  if (g.allSales && !g.allCS) return "Sales Team";
  if (g.allCS && !g.allSales) return "CS Team";
  if (p.hasFYRev) {
    if (p.fyHasSales && !p.fyHasCS) return "Sales Team";
    if (p.fyHasCS && !p.fyHasSales) return "CS Team";
    return "CS and Sales (Transferred this FY)";
  }
  return (g.anySales && g.anyCS) ? "CS & Sales" : "C-Suite";
}

function getFYString(date) {
  const year = date.getFullYear();
  const fyStart = (date.getMonth() >= SUMMARY_CONFIG.FY_START_MONTH) ? year : year - 1;
  return `FY ${fyStart.toString().slice(-2)}/${(fyStart + 1).toString().slice(-2)}`;
}

function columnToLetter(column) {
  let temp, letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
