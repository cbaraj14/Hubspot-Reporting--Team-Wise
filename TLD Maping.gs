// ==== Configurable inputs ====
const INPUT_SHEET_NAME = "All_Payment_Sales_and_CS_Pipeline";            // Name of the input sheet
const OUTPUT_SHEET_NAME = "TLD Mapping with Hubspot";    // Sheet where result will be written
const MAPPING_SHEET_NAME = "TLD Mapping";  // TLD to Account Name mapping

const FREE_PROVIDERS = [
  "gmail", "yahoo", "hotmail", "outlook", "aol", "icloud", "protonmail",
  "proton", "live", "msn", "gmx", "zoho", "tutanota", "pm", "me", "mail",
  "fastmail", "yandex", "rediffmail", "inbox", "qq", "naver", "daum",
  "seznam", "wp", "laposte", "yopmail"
];

// Wrapper function for menu call, hardcoded with sheet names
function runTLDExtractionFromMenu() {
  try {
    const message = runTLDExtraction(INPUT_SHEET_NAME, OUTPUT_SHEET_NAME, MAPPING_SHEET_NAME);
    SpreadsheetApp.getUi().alert(message);
  } catch (error) {
    SpreadsheetApp.getUi().alert("❌ Error: " + error.message);
  }
}

function runTLDExtraction(inputSheetName, outputSheetName, mappingSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(inputSheetName);
  if (!inputSheet) throw new Error(`Sheet "${inputSheetName}" not found.`);

  const data = inputSheet.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1);

  const companyIdx = header.findIndex(h => h.toString().toLowerCase().includes("company") || h.toString().toLowerCase().includes("account"));
  const emailIdx = header.findIndex(h => h.toString().toLowerCase().includes("email"));

  if (companyIdx === -1 || emailIdx === -1) {
    throw new Error("Company or Email column not found.");
  }

  // Build mapping from 'TLD Mapping' sheet: column A = TLD, column B = mapped name
  const mappingSheet = ss.getSheetByName(mappingSheetName);
  if (!mappingSheet) throw new Error(`Sheet "${mappingSheetName}" not found.`);

  const mapData = mappingSheet.getDataRange().getValues();
  const tldMap = new Map();
  for (let i = 1; i < mapData.length; i++) {
    const key = mapData[i][0]?.toString().trim().toLowerCase();
    const val = mapData[i][1]?.toString().trim();
    if (key && val) tldMap.set(key, val);
  }

  const outputMap = new Map(); // Map to ensure no duplicate accounts

  for (let row of rows) {
    const companyRaw = row[companyIdx];
    const emailRaw = row[emailIdx];
    if (!companyRaw || !emailRaw) continue;

    const company = companyRaw.toString().trim();
    if (outputMap.has(company)) continue; // Skip duplicates

    const emails = extractEmails(emailRaw.toString());

    let fullDomain = null;
    for (let email of emails) {
      const domain = email.split("@")[1].toLowerCase();
      const sld = getSLD(domain);
      if (!FREE_PROVIDERS.includes(sld)) {
        fullDomain = domain;
        break;
      }
    }

    if (fullDomain) {
      const mappedAccount = tldMap.has(fullDomain)
        ? tldMap.get(fullDomain)
        : `(Unmapped) ${company}`;  // ← change applied here
      outputMap.set(company, [fullDomain, mappedAccount]);
    }
  }

  // Output headers and values
  const output = [["Hubspot Account Name", "TLD", "Billing Account Name"]];
  for (let [company, [domain, mapped]] of outputMap.entries()) {
    output.push([company, domain, mapped]);
  }

  let outSheet = ss.getSheetByName(outputSheetName);
  if (outSheet) {
    outSheet.clear();
  } else {
    outSheet = ss.insertSheet(outputSheetName);
  }
  outSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  return `✅ Extraction complete. ${output.length - 1} unique rows written to "${outputSheetName}".`;
}

// Utility: extract all emails from text
function extractEmails(text) {
  const regex = /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi;
  return text.match(regex) || [];
}

// Utility: get second-level domain
function getSLD(domain) {
  const parts = domain.split(".");
  return parts.length >= 2 ? parts[parts.length - 2] : "";
}
