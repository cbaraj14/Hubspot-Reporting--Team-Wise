function remapCompanyNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const summarySheet = ss.getSheetByName("Summary");
  const salesSheet = ss.getSheetByName("Sales_team_only");
  const mappingSheet = ss.getSheetByName("TLD Mapping with Hubspot");

  if (!summarySheet || !salesSheet || !mappingSheet) {
    throw new Error("One or more required sheets are missing: 'Summary', 'Sales_team_only', or 'TLD Mapping with Hubspot'.");
  }

  // === Step 1: Load mapping from Hubspot === //
  const mappingData = mappingSheet.getRange(2, 1, mappingSheet.getLastRow() - 1, 3).getValues(); // A:C
  const nameMap = {};

  mappingData.forEach(([hubspotName, , billingName]) => {
    if (typeof hubspotName === "string" && typeof billingName === "string") {
      nameMap[hubspotName.trim().toLowerCase()] = billingName.trim();
    }
  });

  // === Step 2: Define helper to remap company names on a given sheet === //
  function processSheet(sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const companyRange = sheet.getRange(2, 1, lastRow - 1, 1); // Column A, skip header
    const companyValues = companyRange.getValues();

    const updatedValues = companyValues.map(row => {
      const originalName = String(row[0] || "").trim();
      const lowerName = originalName.toLowerCase();

      if (nameMap[lowerName]) {
        // Replace with mapped billing account name
        return [nameMap[lowerName]];
      } else if (!originalName.startsWith("(Unmapped)")) {
        // Prefix only if not already prefixed
        return [`(Unmapped) ${originalName}`];
      } else {
        return [originalName]; // Already prefixed
      }
    });

    companyRange.setValues(updatedValues);
  }

  // === Step 3: Apply mapping to both sheets === //
  processSheet(summarySheet);
  processSheet(salesSheet);
}
