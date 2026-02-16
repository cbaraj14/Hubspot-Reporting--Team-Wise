/**
 * ============================================================================
 * 1) CONFIGURATION for Sales Team Report
 * ============================================================================
 */
const SALES_REPORT_CONFIG = {
  SOURCE_SHEET: "TEMP_DATA",
  CONFIG_SHEET: "Config_sheet",
  STOP_LIST_SHEET: "Company_Stop_list",
  REPORT_SHEET: "Sales Team Report",
  DEFAULT_PIPELINE: "Payment Pipeline", // Fallback if E18 is empty

  // Config Cells
  START_DATE_CELL: "E8",
  END_DATE_CELL: "E9",
  REPORT_DATE_CELL: "E6", // Reference for current month
  FILTER_C_CELL: "E12", // Revenue This FY (TRUE/FALSE)
  FILTER_D_CELL: "E13", // Aggregation Level (Deal Level/Company Level)
  FILTER_E_CELL: "E14", // New Clients Only (New only/All)
  FILTER_F_CELL: "E15", // Payment Threshold for OLD OTP
  FILTER_G_CELL: "E16", // Revenue Growth Check (TRUE/FALSE)
  FILTER_H_CELL: "E17", // Months Window for Transferred/Forecast
  TARGET_PIPELINE_CELL: "E18", // Dynamic Pipeline Cell

  // Source Column Indices (0-based)
  COL_COMPANY: 0,     // "Associated company names"
  COL_AMOUNT: 6,
  COL_PIPELINE: 7,    // Column H
  COL_CLOSE_DATE: 8,
  COL_DEAL_NAME: 11,
  COL_DEAL_TYPE: 15,  // Column P
  COL_CLIENT_AGE: 16, // Column Q
  COL_FIRST_REV_MONTH: 17, // Column R (Date)
  COL_FIRST_FY: 18,   // Column S
  COL_IS_SALES_MEMBER: 20, // Column U (Is Sales Team Member)
  COL_HAS_REV_FY: 21,      // Column V (Has Revenue this FY) - MANDATORY
  COL_IS_CS_MEMBER: 23,    // Column X (Is CS Team Member)
  COL_HAS_SALES_TEAM: 25   // Column Z (Company has Sales Team)
};

const SALES_EXTRA_COLUMNS = [
  { header: "Revenue Trend", isSparkline: true },
  { header: "Revenue Type", index: 15 },
  { header: "Client Age", index: 16 },
  { header: "First Revenue Fiscal Year", index: 18 },
  { header: "POC Team", isVirtual: true }
];

const SALES_TRAILING_COLUMNS = [
  "Realized Revenue (FY) Source",
  "Forecasted Revenue This FY",
  "Total Revenue for the period (Including forecast)"
];

/**
 * ============================================================================
 * 2) MAIN EXECUTION
 * ============================================================================
 */
function salesGenerateSalesTeamReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SALES_REPORT_CONFIG.SOURCE_SHEET);
  const configSheet = ss.getSheetByName(SALES_REPORT_CONFIG.CONFIG_SHEET);
  const stopSheet = ss.getSheetByName(SALES_REPORT_CONFIG.STOP_LIST_SHEET);

  if (!sourceSheet || !configSheet) {
    throw new Error("Source or Config sheet missing.");
  }

  // Load Configurations
  const startDate = new Date(configSheet.getRange(SALES_REPORT_CONFIG.START_DATE_CELL).getValue());
  const endDate = new Date(configSheet.getRange(SALES_REPORT_CONFIG.END_DATE_CELL).getValue());
  const rawReportDate = new Date(configSheet.getRange(SALES_REPORT_CONFIG.REPORT_DATE_CELL).getValue());
  const reportDate = new Date(rawReportDate.getFullYear(), rawReportDate.getMonth(), 1);
  const reportFY = salesGetFYString(endDate);

  // Dynamic Pipeline Logic: Pull from E18, fallback to "Payment Pipeline" if empty
  const targetPipelineRaw = configSheet.getRange(SALES_REPORT_CONFIG.TARGET_PIPELINE_CELL).getValue();
  const targetPipeline = targetPipelineRaw ? targetPipelineRaw.toString().trim() : SALES_REPORT_CONFIG.DEFAULT_PIPELINE;
  
  // FY Boundaries for Forecast logic
  let fyStartYear = reportDate.getMonth() >= 6 ? reportDate.getFullYear() : reportDate.getFullYear() - 1;
  const currentFYStart = new Date(fyStartYear, 6, 1);
  const currentFYEnd = new Date(fyStartYear + 1, 5, 30);

  const filterC_val = configSheet.getRange(SALES_REPORT_CONFIG.FILTER_C_CELL).getValue();
  const filterD_val = configSheet.getRange(SALES_REPORT_CONFIG.FILTER_D_CELL).getValue();
  const filterE_val = configSheet.getRange(SALES_REPORT_CONFIG.FILTER_E_CELL).getValue();
  const filterF_threshold = configSheet.getRange(SALES_REPORT_CONFIG.FILTER_F_CELL).getValue() || 0;
  const filterG_val = configSheet.getRange(SALES_REPORT_CONFIG.FILTER_G_CELL).getValue();
  const filterH_months = configSheet.getRange(SALES_REPORT_CONFIG.FILTER_H_CELL).getValue() || 12;

  const stopList = stopSheet ? stopSheet.getDataRange().getValues().map(r => r[0].toString().toLowerCase()) : [];
  const rawData = sourceSheet.getDataRange().getValues();
  const rows = rawData.slice(1);
  
  const companyGroups = {};
  rows.forEach(row => {
    const co = row[SALES_REPORT_CONFIG.COL_COMPANY];
    if (!co) return;

    // PIPELINE FILTER: Only include deals in the Target Pipeline (from E18 or Default)
    if (row[SALES_REPORT_CONFIG.COL_PIPELINE] !== targetPipeline) return;

    if (!companyGroups[co]) companyGroups[co] = [];
    companyGroups[co].push(row);
  });

  const pivotMap = {};
  const allMonths = new Set();

  // Populate all months range for headers
  let tempH = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  while (tempH <= endDate) {
    allMonths.add(Utilities.formatDate(tempH, Session.getScriptTimeZone(), "yyyy-MMM"));
    tempH.setMonth(tempH.getMonth() + 1);
  }

  // B) APPLY FILTERS & AGGREGATE
  for (const companyName in companyGroups) {
    const deals = companyGroups[companyName];

    // --- MANDATORY FILTERS ---
    if (stopList.includes(companyName.toLowerCase())) continue;
    if (deals[0][SALES_REPORT_CONFIG.COL_HAS_SALES_TEAM] !== true) continue;
    
    // Mandatory check for "Has Revenue this FY" (Column V)
    if (!deals.some(d => d[SALES_REPORT_CONFIG.COL_HAS_REV_FY] === true)) continue;

    // --- CONDITIONAL FILTERS ---
    if (filterC_val === true && !deals.some(d => d[SALES_REPORT_CONFIG.COL_HAS_REV_FY] === true)) continue;
    if (filterE_val === "New only" && !deals.some(d => d[SALES_REPORT_CONFIG.COL_CLIENT_AGE] === "New")) continue;

    const isOldOTP = deals.some(d => d[SALES_REPORT_CONFIG.COL_CLIENT_AGE] === "Old" && d[SALES_REPORT_CONFIG.COL_DEAL_TYPE] === "OTP");
    if (isOldOTP) {
        const fyPayments = deals.filter(d => salesGetFYString(new Date(d[SALES_REPORT_CONFIG.COL_CLOSE_DATE])) === reportFY).length;
        if (fyPayments > filterF_threshold) continue;
    }

    if (filterG_val === true) {
        const isOldRec = deals.some(d => d[SALES_REPORT_CONFIG.COL_CLIENT_AGE] === "Old" && (d[SALES_REPORT_CONFIG.COL_DEAL_TYPE] === "Recurring" || d[SALES_REPORT_CONFIG.COL_DEAL_TYPE] === "R-OTP"));
        if (isOldRec && !salesHasRevenueGrowth(deals)) continue;
    }

    // Reference Column R for First Revenue Month
    const firstRevDateRaw = deals[0][SALES_REPORT_CONFIG.COL_FIRST_REV_MONTH];
    const firstRevDate = firstRevDateRaw instanceof Date ? firstRevDateRaw : null;
    
    let filteredDeals = deals;
    const isRecOrROTP = deals.some(d => d[SALES_REPORT_CONFIG.COL_DEAL_TYPE] === "Recurring" || d[SALES_REPORT_CONFIG.COL_DEAL_TYPE] === "R-OTP");
    
    if (isRecOrROTP) {
        filteredDeals = deals.filter(d => {
            if (d[SALES_REPORT_CONFIG.COL_IS_CS_MEMBER] !== true) return true;
            if (!firstRevDate) return false;
            const closeDate = new Date(d[SALES_REPORT_CONFIG.COL_CLOSE_DATE]);
            const diffMonths = (closeDate.getFullYear() - firstRevDate.getFullYear()) * 12 + (closeDate.getMonth() - firstRevDate.getMonth());
            return diffMonths >= 0 && diffMonths <= filterH_months;
        });
    }

    filteredDeals.forEach(row => {
      if (filterD_val === "Deal Level" && row[SALES_REPORT_CONFIG.COL_IS_SALES_MEMBER] !== true) return;

      const closeDate = new Date(row[SALES_REPORT_CONFIG.COL_CLOSE_DATE]);
      if (closeDate < startDate || closeDate > endDate) return;

      const monthKey = Utilities.formatDate(closeDate, Session.getScriptTimeZone(), "yyyy-MMM");
      if (!pivotMap[companyName]) {
        pivotMap[companyName] = {
          name: companyName,
          tempRow: row,
          monthlyValues: {},
          realizedFYSource: 0,
          firstRevDate: firstRevDate,
          firstFY: row[SALES_REPORT_CONFIG.COL_FIRST_FY] // Used for sorting
        };
      }

      const p = pivotMap[companyName];
      const amount = parseFloat(row[SALES_REPORT_CONFIG.COL_AMOUNT]) || 0;
      p.monthlyValues[monthKey] = (p.monthlyValues[monthKey] || 0) + amount;
      if (salesGetFYString(closeDate) === reportFY) {
        p.realizedFYSource += amount;
      }
    });
  }

  salesRenderReport(ss, pivotMap, allMonths, reportFY, reportDate, currentFYStart, currentFYEnd, filterH_months);
}

/**
 * Helper to check for instances of monthly revenue increase
 */
function salesHasRevenueGrowth(deals) {
    const monthlyTotals = {};
    deals.forEach(d => {
        const m = Utilities.formatDate(new Date(d[SALES_REPORT_CONFIG.COL_CLOSE_DATE]), Session.getScriptTimeZone(), "yyyy-MM");
        monthlyTotals[m] = (monthlyTotals[m] || 0) + (parseFloat(d[SALES_REPORT_CONFIG.COL_AMOUNT]) || 0);
    });
    const sortedMonths = Object.keys(monthlyTotals).sort();
    for (let i = 1; i < sortedMonths.length; i++) {
        if (monthlyTotals[sortedMonths[i]] > monthlyTotals[sortedMonths[i-1]]) return true;
    }
    return false;
}

/**
 * Main Rendering function matching CS report UI/Logic
 */
function salesRenderReport(ss, pivotMap, allMonthsSet, reportFY, reportDate, fyStart, fyEnd, hMonths) {
  const sortedMonthsKeys = Array.from(allMonthsSet).sort((a,b) => new Date(a) - new Date(b));
  const dynamicHeaders = SALES_EXTRA_COLUMNS.map(c => c.header);
  const headers = ["Company Name", ...dynamicHeaders, ...sortedMonthsKeys, ...SALES_TRAILING_COLUMNS];
  
  const finalOutput = [headers];
  
  // SORTING LOGIC: Newest First FY (desc), then Company Name (asc)
  const sortedCompanies = Object.values(pivotMap).sort((a, b) => {
    const fyA = String(a.firstFY || "");
    const fyB = String(b.firstFY || "");
    if (fyA !== fyB) return fyB.localeCompare(fyA);
    return a.name.localeCompare(b.name);
  });

  sortedCompanies.forEach(p => {
    const row = [p.name];
    
    // Metadata
    SALES_EXTRA_COLUMNS.forEach(col => {
      if (col.isSparkline) row.push("");
      else if (col.isVirtual) row.push("Sales Team");
      else row.push(p.tempRow[col.index]);
    });

    // Baseline for forecast
    const curMonthKey = Utilities.formatDate(reportDate, Session.getScriptTimeZone(), "yyyy-MMM");
    const prevDate = new Date(reportDate);
    prevDate.setMonth(prevDate.getMonth() - 1);
    const prevMonthKey = Utilities.formatDate(prevDate, Session.getScriptTimeZone(), "yyyy-MMM");
    
    const curVal = p.monthlyValues[curMonthKey] || 0;
    const prevVal = p.monthlyValues[prevMonthKey] || 0;
    const forecastBaseline = curVal > 0 ? curVal : prevVal;

    const revType = (p.tempRow[SALES_REPORT_CONFIG.COL_DEAL_TYPE] || "").toLowerCase();
    const isRecurring = revType.includes("recurring") || revType.includes("r-otp");

    let forecastedFYTotal = 0;
    let periodTotalInclForecast = 0;

    sortedMonthsKeys.forEach(mKey => {
      let val = p.monthlyValues[mKey] || 0;
      const mDate = new Date(mKey);
      const isHistorical = mDate < reportDate;

      // Apply Forecast Logic: 
      if (!isHistorical && val === 0 && isRecurring) {
        const withinFY = mDate <= fyEnd;
        let withinWindow = true;
        if (p.firstRevDate) {
          const diffMonths = (mDate.getFullYear() - p.firstRevDate.getFullYear()) * 12 + (mDate.getMonth() - p.firstRevDate.getMonth());
          withinWindow = diffMonths <= hMonths;
        }
        if (withinWindow && withinFY) val = forecastBaseline;
      }

      row.push(val);
      if (mDate >= fyStart && mDate <= fyEnd) forecastedFYTotal += val;
      periodTotalInclForecast += val;
    });

    row.push(p.realizedFYSource, forecastedFYTotal, periodTotalInclForecast);
    finalOutput.push(row);
  });

  let sheet = ss.getSheetByName(SALES_REPORT_CONFIG.REPORT_SHEET);
  if (!sheet) sheet = ss.insertSheet(SALES_REPORT_CONFIG.REPORT_SHEET);
  
  sheet.clear();
  if (finalOutput.length > 1) {
    const lastRow = finalOutput.length;
    const lastCol = finalOutput[0].length;
    const staticColCount = 1 + SALES_EXTRA_COLUMNS.length;
    const monthStartIdx = staticColCount + 1;

    sheet.getRange(1, 1, lastRow, lastCol).setValues(finalOutput);
    
    // Sparkline Logic
    const sparkIdx = SALES_EXTRA_COLUMNS.findIndex(c => c.isSparkline);
    if (sparkIdx !== -1) {
      const sparkCol = sparkIdx + 2;
      for (let i = 2; i <= lastRow; i++) {
        const start = salesColumnToLetter(monthStartIdx) + i;
        const end = salesColumnToLetter(monthStartIdx + sortedMonthsKeys.length - 1) + i;
        sheet.getRange(i, sparkCol).setFormula(`=SPARKLINE(${start}:${end}, {"charttype","line";"linewidth",2;"linecolor","#5f6368"})`);
      }
    }

    // Totals Row
    const totalRowIdx = lastRow + 1;
    sheet.getRange(totalRowIdx, 1).setValue("TOTALS").setFontWeight("bold");
    const totalCols = [...Array(sortedMonthsKeys.length).keys()].map(x => x + monthStartIdx).concat([lastCol - 2, lastCol - 1, lastCol]);
    totalCols.forEach(c => {
      sheet.getRange(totalRowIdx, c).setFormula(`=SUBTOTAL(9, ${salesColumnToLetter(c)}2:${salesColumnToLetter(c)}${lastRow})`).setFontWeight("bold");
    });

    // Final Styling
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(2);
    sheet.getRange(1, 1, 1, lastCol).setBackground("#444444").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange(2, monthStartIdx, totalRowIdx, sortedMonthsKeys.length + 3).setNumberFormat('[=0]"-" ; $#,##0');
    
    // Trend Coloring
    sortedMonthsKeys.forEach((mKey, i) => {
      const colIdx = monthStartIdx + i;
      const mDate = new Date(mKey);
      const range = sheet.getRange(2, colIdx, lastRow - 1, 1);
      if (mDate >= fyStart && mDate <= reportDate) {
        range.setFontColor("black");
        if (i > 0) {
          const vals = range.getValues();
          const prevVals = sheet.getRange(2, colIdx - 1, lastRow - 1, 1).getValues();
          const colors = vals.map((v, idx) => {
            const curr = parseFloat(v[0]) || 0, prev = parseFloat(prevVals[idx][0]) || 0;
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
  } else {
    SpreadsheetApp.getUi().alert("No data matched the Sales Team criteria.");
  }
}

function salesGetFYString(date) {
  const year = date.getFullYear();
  const fyStart = (date.getMonth() >= 6) ? year : year - 1;
  return `FY ${fyStart.toString().slice(-2)}/${(fyStart + 1).toString().slice(-2)}`;
}

function salesColumnToLetter(column) {
  let temp, letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
