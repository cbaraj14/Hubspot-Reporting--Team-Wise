# Contributing Guide

Thank you for contributing to the HubSpot Deals to Google Sheets Importer! This guide will help you make changes safely and effectively.

---

## 🎯 Quick Guidelines

1. **Always test in a copy** of the spreadsheet first
2. **Document all changes** in code comments and this README
3. **Follow existing patterns** for consistency
4. **Update README.md** if adding features
5. **Test thoroughly** before deploying to production

---

## 📋 Prerequisites for Contributors

### Required Knowledge
- ✅ Basic JavaScript (ES6+)
- ✅ Google Apps Script fundamentals
- ✅ HubSpot CRM basics
- ✅ Google Sheets functions

### Recommended Knowledge
- 🔧 REST APIs and HTTP requests
- 🔧 Data processing and filtering
- 🔧 Google Sheets advanced formulas

### Tools You'll Need
- Google account with Sheets access
- Text editor (VS Code recommended)
- Apps Script editor (built into Google Sheets)
- HubSpot account for testing (optional)

---

## 🏗️ Development Setup

### 1. Create a Development Copy

```
1. Open the production Google Sheet
2. File → Make a copy
3. Rename: "[DEV] HubSpot Integration - Your Name"
4. Work in this copy (never in production!)
```

### 2. Set Up Your Test Environment

**Create test data:**
- Use a separate HubSpot sandbox account if available
- Or use production data but with a short date range
- Limit to 100-500 deals for fast testing

**Configure test settings:**
- Update `Config_sheet` E8/E9 with narrow date range
- Point to test team member sheets
- Use test exclusion lists

### 3. Enable Logging

Add logging to your code:

```javascript
// At the start of your function
console.log('=== Function Started ===');
const startTime = new Date();

// Throughout your code
console.log('Step 1 complete:', someVariable);

// At the end
const endTime = new Date();
console.log('=== Function Complete ===');
console.log('Time taken:', (endTime - startTime) / 1000, 'seconds');
```

View logs:
- **Extensions** → **Apps Script** → **Executions**

---

## 🛠️ Making Changes

### Code Style

#### Naming Conventions

```javascript
// Constants - UPPERCASE_WITH_UNDERSCORES
const MAX_RETRIES = 3;
const API_BASE_URL = 'https://api.hubapi.com';

// Config objects - CAPS with descriptive suffix
const SUMMARY_CONFIG = { ... };
const SALES_REPORT_CONFIG = { ... };

// Functions - camelCase, descriptive verb + noun
function generateSummaryReport() { }
function fetchDealsFromHubSpot() { }
function validateConfigSettings() { }

// Variables - camelCase, descriptive
const reportStartDate = new Date();
const companyPivotMap = {};
const hasValidToken = false;

// Private helpers - prefix with underscore (optional)
function _calculateFiscalYear(date) { }
```

#### Function Structure

```javascript
/**
 * Brief description of what the function does
 * 
 * @param {Object} config - Configuration object
 * @param {Date} startDate - Start date for filtering
 * @returns {Array} Array of processed deals
 */
function processDeals(config, startDate) {
  // 1. Input validation
  if (!config || !startDate) {
    throw new Error('Missing required parameters');
  }
  
  // 2. Setup
  const results = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 3. Main logic
  try {
    // ... processing code
    
  } catch (error) {
    // 4. Error handling
    console.error('Error in processDeals:', error);
    throw error;
  }
  
  // 5. Return
  return results;
}
```

#### Comments

```javascript
// ✅ GOOD - Explains WHY, not WHAT
// Use last modified date to avoid re-processing unchanged deals
const lastSync = getLastSyncDate();

// ❌ BAD - Just repeats the code
// Get last sync date
const lastSync = getLastSyncDate();

// ✅ GOOD - Complex business logic
// Per finance team: Recurring = 8+ months in one FY OR keyword in name
if (monthCount >= 8 || dealName.includes('subscription')) {
  revenueType = 'Recurring';
}

// ❌ BAD - Obvious code
// Set revenue type to recurring
revenueType = 'Recurring';
```

#### Error Handling

```javascript
// Always use try-catch for external API calls
try {
  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  return data;
  
} catch (error) {
  // Log the error
  console.error('API call failed:', {
    url: url,
    error: error.message,
    timestamp: new Date()
  });
  
  // User-friendly alert
  SpreadsheetApp.getUi().alert(
    'Error fetching data from HubSpot. Please check logs.'
  );
  
  // Re-throw or return null
  throw error;
}
```

---

## 🧪 Testing Your Changes

### Testing Checklist

Before deploying, verify:

#### Functionality Tests
- [ ] Function executes without errors
- [ ] Output matches expected format
- [ ] Data accuracy (spot check 10-20 rows)
- [ ] Edge cases handled (empty data, missing fields, etc.)

#### Performance Tests
- [ ] Execution time acceptable (< 6 minutes)
- [ ] No timeout errors
- [ ] Memory usage within limits
- [ ] API rate limits respected

#### UI/UX Tests
- [ ] Custom menu appears correctly
- [ ] Success/error messages are clear
- [ ] Sheet formatting is correct
- [ ] Frozen rows/columns work
- [ ] Filters apply correctly

#### Integration Tests
- [ ] Doesn't break existing reports
- [ ] Other menu functions still work
- [ ] TEMP_DATA sheet updates correctly
- [ ] Configuration cells work as expected

### Test Data Scenarios

Test with these scenarios:

1. **Empty Dataset**
   - Set date range with no deals
   - Should complete without errors
   - Should show "No data" message

2. **Single Row**
   - Limit to 1 deal
   - Verify formatting applies
   - Check subtotal formulas work

3. **Large Dataset**
   - 1,000+ deals
   - Monitor execution time
   - Check for timeouts

4. **Missing Data**
   - Deals without companies
   - Deals without contacts
   - Invalid dates
   - Null amounts

5. **Edge Cases**
   - Report Date = today
   - Date range spanning 3+ years
   - All deals from one owner
   - Companies with special characters in name

---

## 📝 Common Tasks

### Adding a New Filter to Sales Team Report

1. **Add configuration cell**

```javascript
// In SALES_REPORT_CONFIG
FILTER_NEW_CELL: "E19",  // Add to next available cell
```

2. **Read the filter value**

```javascript
const filterNewValue = configSheet
  .getRange(SALES_REPORT_CONFIG.FILTER_NEW_CELL)
  .getValue();
```

3. **Apply the filter logic**

```javascript
function applyNewFilter(companyData, filterValue) {
  if (!filterValue) return true;  // Skip if not enabled
  
  // Your filter logic here
  return companyData.someProperty === filterValue;
}
```

4. **Integrate into main filter chain**

```javascript
// In salesGenerateSalesTeamReport()
if (!applyNewFilter(company, filterNewValue)) {
  continue;  // Skip this company
}
```

5. **Update README.md** with new filter documentation

### Adding a New Column to TEMP_DATA

1. **Update OUTPUT_HEADERS in Config**

```javascript
const OUTPUT_HEADERS = [
  "Company Name", "Deal Stage", ...,
  "Your New Column"  // Add here
];
```

2. **Update Temp Sheet.gs processing**

```javascript
// In tempUpdateTempSheet(), add to finalOutput.push()
finalOutput.push([
  row[11], row[3], ...,
  calculateNewValue(row)  // Your logic
]);
```

3. **Update report column indices**

```javascript
// In report CONFIG objects
COL_NEW_FIELD: 26,  // Next available index
```

4. **Test thoroughly** - column order is critical!

### Adding a New Report

1. **Create new .gs file**

```javascript
// NewReport.gs
const NEW_REPORT_CONFIG = {
  SOURCE_SHEET: "TEMP_DATA",
  CONFIG_SHEET: "Config_sheet",
  REPORT_SHEET: "New Report",
  // ... other config
};

function generateNewReport() {
  // Your report logic
}
```

2. **Add to Menu.gs**

```javascript
ui.createMenu('HubSpot Sync')
  // ... existing items
  .addItem('8. Generate New Report', 'generateNewReport')
  .addToUi();
```

3. **Update README.md** with usage instructions

### Modifying Pipeline Mappings

1. **Update CONFIG in Config file**

```javascript
PIPELINE_MAP: {
  '70710959': 'Sales Pipeline',
  '679755281': 'CS Pipeline',
  'your-new-id': 'Your Pipeline Name',  // Add here
  'default': 'Payment Pipeline'
}
```

2. **Update SOURCE_SHEETS array**

```javascript
const SOURCE_SHEETS = [
  PIPELINE_MAP['70710959'],
  PIPELINE_MAP['679755281'],
  PIPELINE_MAP['your-new-id'],  // Add here
  PIPELINE_MAP.default
];
```

3. **Test sync with new pipeline**

---

## 🚨 Common Pitfalls

### 1. Column Index Off-by-One Errors

**Problem:** Arrays are 0-indexed, but Sheets are 1-indexed

```javascript
// ❌ WRONG
const dealId = row[8];  // But you need column I (9th column)

// ✅ CORRECT
const dealId = row[8];  // Column I is index 8 (0-based)
```

**Tip:** Always count from 0 for arrays, 1 for Sheets

### 2. Date Comparison Issues

**Problem:** JavaScript Date objects are tricky

```javascript
// ❌ WRONG
if (date1 == date2) { }  // Compares objects, not values

// ✅ CORRECT
if (date1.getTime() === date2.getTime()) { }
// OR
if (date1.toDateString() === date2.toDateString()) { }
```

### 3. API Rate Limiting

**Problem:** Too many requests too fast

```javascript
// ❌ WRONG - No delay between calls
for (let i = 0; i < 1000; i++) {
  UrlFetchApp.fetch(url);  // Will hit rate limit!
}

// ✅ CORRECT - Add delays
for (let i = 0; i < 1000; i++) {
  if (i % 100 === 0 && i > 0) {
    Utilities.sleep(300);  // Pause every 100 calls
  }
  UrlFetchApp.fetch(url);
}
```

### 4. Sheet Name Case Sensitivity

**Problem:** Sheet names are case-sensitive

```javascript
// ❌ WRONG
const sheet = ss.getSheetByName('summary report');  // Won't find it!

// ✅ CORRECT
const sheet = ss.getSheetByName('Summary Report');  // Exact match
```

### 5. Modifying Config During Execution

**Problem:** Changing CONFIG while code is running

```javascript
// ❌ WRONG
CONFIG.BATCH_SIZE = 5000;  // Doesn't affect already-running code

// ✅ CORRECT
// Stop execution, change CONFIG, save, re-run
```

---

## 📤 Submitting Changes

### Pre-Submission Checklist

- [ ] Code tested in dev copy
- [ ] All tests pass (see Testing Checklist)
- [ ] Code follows style guide
- [ ] Comments added for complex logic
- [ ] README.md updated (if applicable)
- [ ] CHANGELOG.md updated
- [ ] No sensitive data (tokens, real company names)
- [ ] No console.log() statements (use Logger.log())

### Deployment Process

1. **Review Your Changes**
   ```
   - Compare dev copy to production
   - Note all modified functions
   - Document any breaking changes
   ```

2. **Create Backup**
   ```
   - File → Make a copy of production sheet
   - Name: "[BACKUP] HubSpot Integration - [Date]"
   - Store in "Backups" folder
   ```

3. **Deploy to Production**
   ```
   - Open production Google Sheet
   - Extensions → Apps Script
   - Copy modified code from dev copy
   - Save (Ctrl+S / Cmd+S)
   ```

4. **Test in Production**
   ```
   - Run with small date range first
   - Verify outputs match dev copy
   - Check no errors in execution logs
   ```

5. **Announce Changes**
   ```
   - Email team with change summary
   - Note any new features/changes
   - Document known issues (if any)
   ```

---

## 🐛 Debugging Tips

### Enable Verbose Logging

```javascript
const DEBUG = true;  // Set to false in production

function myFunction() {
  if (DEBUG) console.log('Starting myFunction');
  
  // Your code
  
  if (DEBUG) console.log('Finished myFunction');
}
```

### Inspect Data at Breakpoints

```javascript
// Add temporary inspection code
console.log('=== DATA INSPECTION ===');
console.log('Variable:', JSON.stringify(myVariable, null, 2));
console.log('Type:', typeof myVariable);
console.log('Length:', myVariable?.length);
console.log('======================');
```

### Use Apps Script Debugger

1. Open Apps Script editor
2. Set breakpoint (click line number)
3. Click **Debug** (bug icon)
4. Inspect variables in right panel

### Check Execution Logs

```
Extensions → Apps Script → Executions
- Click any execution to see logs
- Filter by status (Success, Error)
- Look for error stack traces
```

---

## 📚 Resources

### Documentation
- [README.md](README.md) - Full system documentation
- [QUICKSTART.md](QUICKSTART.md) - Setup guide
- [CHANGELOG.md](CHANGELOG.md) - Version history

### External Resources
- [Google Apps Script Docs](https://developers.google.com/apps-script)
- [HubSpot API Docs](https://developers.hubspot.com/docs/api/overview)
- [JavaScript MDN Docs](https://developer.mozilla.org/en-US/docs/Web/JavaScript)

### Getting Help
1. Check execution logs first
2. Review README troubleshooting section
3. Search existing code for similar patterns
4. Contact internal development team

---

## 🙏 Thank You!

Your contributions make this project better for everyone. Questions? Reach out to the development team.

Happy coding! 🚀

---

*Last Updated: February 16, 2025*
