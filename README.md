# HubSpot Deals to Google Sheets Importer

**Professional Google Apps Script Integration for HubSpot CRM Data**

---

## 📋 Overview

This Google Apps Script solution provides a robust, non-developer-friendly integration between HubSpot CRM and Google Sheets. It automatically imports deal data from multiple HubSpot pipelines with real-time progress tracking and intelligent date-based filtering, then generates comprehensive revenue reports with forecasting capabilities.

---

## 🎯 Key Features

### Non-Developer Friendly Design
- **Centralized Configuration**: All settings (API tokens, sheet names, date ranges) are consolidated in a single `CONFIG` object
- **No Coding Required**: Users can modify pipeline IDs, date ranges, and sheet preferences without touching any code logic
- **Clear Documentation**: Inline comments explain each configurable parameter

### 📊 Multi-Pipeline Support
Imports deals from three distinct HubSpot pipelines:
- **Sales Pipeline** (ID: 70710959)
- **CS Pipeline** (ID: 679755281)
- **Payment Pipeline** (default - source of truth for revenue)

Automatically maps deal stages to human-readable names:
- Closed Won (Sales Pipeline) - Stage ID: 136833349
- Revenue (Payment Pipeline) - Stage: closedwon
- Closed Won (Payment Pipeline) - Stage ID: 996632637

### 📅 Smart Date Filtering
- Reads custom start date from a configurable Config Sheet (cell E6 by default)
- Falls back to default start date (2023-01-01) if no date is specified
- Ensures only relevant deals are imported based on creation date

### 📋 Comprehensive Data Capture

Each imported deal includes 15 essential fields:

| Column | Data Point | HubSpot Property |
|--------|-----------|------------------|
| A-D | Deal Basics | Deal Name, Amount, Create Date, Deal Stage |
| E | Pipeline | Pipeline Name |
| F-H | Relationships | Owner ID, Company ID, Contact ID |
| I-K | Metadata | Deal ID (hs_object_id), Close Date, Last Modified Date |
| L-O | Enriched Data | Company Name, Contact Name, Owner Name, Contact Email |

### 📈 Real-Time Progress Tracking
- Dedicated progress logging in App Script terminal
- Shows total deals found across all pipelines
- Breaks down deal counts by individual pipeline
- Updates in real-time during import process

### ✨ Automatic Formatting & Organization
- **Auto-Sort**: All sheets automatically sort by "Date Created" (descending) for newest-first viewing
- **Header Styling**: Professional header row formatting applied automatically
- **Frozen Headers**: Top row frozen for easy scrolling through large datasets
- **Clean Layout**: Consistent formatting across all generated sheets

### 🔧 Technical Highlights
- Efficient HubSpot API v3 integration with batch processing
- Pagination handling for large datasets (100 deals per batch)
- Robust error handling and validation
- Rate limiting with smart sleep intervals (300ms between batches)
- Modular, maintainable code structure
- Reusable across different Google Sheets
- Deduplication by Deal ID (keeps most recent)

---

## 🏗️ Architecture

### File Structure

```
├── Config                      # Central configuration object
├── hubspot_sync.gs            # HubSpot data import & sync functions
├── Temp Sheet.gs              # Data processing & classification engine
├── Summary reprot.gs          # Summary Report generator
├── Sales Team Report.gs       # Sales Team Report with filtering
├── CS Team Report.gs          # CS Team Report with forecasting
├── HubSpot Company Mapper.gs  # Company name enrichment
├── TLD Maping.gs              # Domain-based company identification
├── Map TLD in Summary.gs      # TLD mapping for reports
└── Menu.gs                    # Custom menu for Google Sheets UI
```

### Data Flow

```
┌─────────────────────┐
│   HubSpot CRM API   │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│  hubspot_sync.gs    │ ← Imports deals from 3 pipelines
│  - Sales Pipeline   │
│  - CS Pipeline      │
│  - Payment Pipeline │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│ Pipeline Sheets     │ ← Raw deal data (15 columns)
│ - Sales Pipeline    │
│ - CS Pipeline       │
│ - Payment Pipeline  │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│  Temp Sheet.gs      │ ← Classification & enrichment
│  - Deduplication    │
│  - Revenue Type     │
│  - Client Age       │
│  - Team Attribution │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│  TEMP_DATA Sheet    │ ← Hidden cache (25+ columns)
└──────────┬──────────┘
           │
           ├──────────────────┬──────────────────┐
           ▼                  ▼                  ▼
┌──────────────────┐ ┌──────────────────┐ ┌──────────────────┐
│ Summary Report   │ │ Sales Team Report│ │ CS Team Report   │
│ - All companies  │ │ - Filtered view  │ │ - CS + Forecast  │
│ - POC tagging    │ │ - Growth logic   │ │ - FY projections │
└──────────────────┘ └──────────────────┘ └──────────────────┘
```

---

## 🚀 Getting Started

### Prerequisites
1. Access to HubSpot CRM with API access
2. HubSpot Private App token with the following scopes:
   - `crm.objects.deals.read`
   - `crm.objects.companies.read`
   - `crm.objects.contacts.read`
3. Google Sheets document
4. Basic understanding of Google Apps Script

### Initial Setup

#### 1. Configure HubSpot Token

Open the `Config` file and add your HubSpot token:

```javascript
const CONFIG = (() => {
  return {
    HUBSPOT_TOKEN: 'your-hubspot-token-here',
    // ... rest of config
  };
})();
```

#### 2. Create Required Sheets

The script requires these sheets to exist in your Google Sheets document:

**Input/Config Sheets:**
- `Config_sheet` - Configuration parameters
- `Sales_team_Members` - List of Sales team Owner IDs (Column A with header)
- `CS_team_Members` - List of CS team Owner IDs (Column A with header)
- `Company_Stop_list` - Companies to exclude from Sales report (Column A with header)

**Auto-Created Sheets:**
- `Sales Pipeline` - Auto-created during first sync
- `CS Pipeline` - Auto-created during first sync
- `Payment Pipeline` - Auto-created during first sync
- `TEMP_DATA` - Hidden sheet for processed data
- `Summary Report` - Summary report output
- `Sales Team Report` - Sales-filtered report output
- `CS Team Report` - CS-focused report with forecasting

#### 3. Configure Date Parameters

In the `Config_sheet`, set these cells:

| Cell | Parameter | Example | Description |
|------|-----------|---------|-------------|
| E6 | Report Date | 6/30/2025 | Reference date for reporting (defaults to today) |
| E8 | Report Start Date | 1/1/2024 | Beginning of report period |
| E9 | Report End Date | 6/30/2026 | End of report period |
| E12 | Sales: Revenue This FY | TRUE/FALSE | Filter sales companies with FY revenue |
| E13 | Sales: Aggregation Level | Deal Level / Company Level | How to aggregate sales data |
| E14 | Sales: Client Age Filter | New only / All | Include only new clients or all |
| E15 | Sales: Payment Threshold | 5 | Max payments for OLD OTP companies |
| E16 | Sales: Growth Check | TRUE/FALSE | Require revenue growth for OLD recurring |
| E17 | Sales: Transferred Months | 12 | Months window for transferred accounts |

---

## 📚 Usage Guide

### Custom Menu

After opening your Google Sheet, you'll see a custom **"HubSpot Sync"** menu with these options:

#### 1. Sync New Deals
**Function:** `syncHubSpotDeals()`

Fetches all deals modified since the last sync from HubSpot. Filters by the configured stage IDs (Closed Won stages only).

**Process:**
- Reads last sync timestamp from Sheet A1
- Queries HubSpot API for deals modified since that date
- Batches deals by pipeline (100 per batch)
- Appends new deals to respective pipeline sheets
- Respects rate limits with 300ms delays

**When to Use:** 
- Daily/weekly to keep data current
- After making changes in HubSpot

#### 2. Sync Missing Companies
**Function:** `syncMissingCompanies()`

Fills in missing company associations for deals.

**Logic:**
- Scans all pipeline sheets for deals without Company ID (Column G)
- Calls HubSpot Batch Associations API
- If no company found, creates fallback name:
  - `Individual-(Contact Name)` if contact exists
  - Email address if no contact name
  
**When to Use:**
- After initial import
- When company associations are missing

#### 3. Sync Missing Contacts
**Function:** `syncMissingContacts()`

Fills in missing contact associations for deals.

**Logic:**
- Scans all pipeline sheets for deals without Contact ID (Column H)
- Calls HubSpot Batch Associations API
- Sets "N/A" if no contact found

**When to Use:**
- After initial import
- When contact associations are missing

#### 4. Update ID to Names
**Function:** `syncAllPipelineData()`

Enriches deals with human-readable names from IDs.

**Enriches:**
- Company Name (from Company ID)
- Contact Name + Email (from Contact ID)
- Owner Name (from Owner ID)

**When to Use:**
- After running sync functions
- To refresh names after changes in HubSpot

#### 5. Generate Summary
**Function:** `generateSummaryReport()`

Creates a comprehensive summary report with all companies.

**Features:**
- Monthly revenue breakdown
- Revenue Type classification (Recurring/R-OTP/OTP)
- Client Age classification (New/Old/Future)
- POC Team assignment
- Sparkline visualizations
- Realized vs Total revenue

**When to Use:**
- For executive dashboards
- Monthly/quarterly reporting
- Full portfolio analysis

#### 6. Run Sales Team Report
**Function:** `salesGenerateSalesTeamReport()`

Creates a filtered report for Sales team with advanced business rules.

**Filters:**
- Excludes companies from Stop List
- Requires Sales team membership
- Optional: Revenue this FY
- Optional: New clients only
- Growth requirements for OLD Recurring
- Payment thresholds for OLD OTP

**When to Use:**
- Sales pipeline review
- Commission calculations
- New business tracking

#### 7. Run CS Team Report
**Function:** `generateCSReport()`

Creates a CS-focused report with revenue forecasting.

**Features:**
- Filters by CS team ownership
- Forecasts recurring revenue through FY end
- Uses last payment amount for projections
- Separate Realized vs Forecasted columns
- No 12-month forecast limit (unlike Sales)

**When to Use:**
- Customer success planning
- Renewal forecasting
- Account health monitoring

---

## 🧮 Business Logic

### Fiscal Year Definition

**Start Date:** July 1  
**End Date:** June 30  
**Format:** FY YY/YY (e.g., FY 25/26 = July 1, 2025 - June 30, 2026)

```javascript
FISCAL_YEAR_START_MONTH: 6  // 0 = January, 6 = July
```

### Revenue Type Classification

Companies are classified into three revenue types based on payment patterns:

#### 1. Recurring Revenue
**Criteria (ANY must be true):**
- Deal name contains "data-services-subscription" OR "gold monthly" (case-insensitive)
- Company has ≥8 unique paid months in ANY single fiscal year

**Example:**
```
Company A: Monthly payments Jan-Dec 2024 → Recurring ✓
Company B: Deal name "Gold Monthly Plan" → Recurring ✓
```

#### 2. R-OTP (Repeated One-Time Payments)
**Criteria:**
- NOT Recurring
- Total unique paid months across all history > 5

**Example:**
```
Company C: 7 payments spread over 3 years → R-OTP ✓
```

#### 3. OTP (One-Time Payment)
**Criteria:**
- Does not meet Recurring or R-OTP criteria
- Typically 1-5 payments total

**Example:**
```
Company D: Single project payment → OTP ✓
```

### Client Age Classification

Based on the company's **first payment date**:

#### New Client
**Criteria:**
- First payment date ≥ July 1 of current fiscal year
- First payment date ≤ Report Date

**Example:**
```
Report Date: 12/31/2024
Current FY: FY 24/25 (July 1, 2024 - June 30, 2025)
First Payment: 8/15/2024 → New ✓
```

#### Old Client
**Criteria:**
- First payment date < July 1 of current fiscal year

**Example:**
```
Current FY: FY 24/25
First Payment: 3/10/2024 → Old ✓
```

#### Future Client
**Criteria:**
- First payment date > Report Date

**Example:**
```
Report Date: 12/31/2024
First Payment: 2/1/2025 → Future ✓
```

### POC Team Classification

The system determines the Point of Contact (POC) team for each company using a hierarchical approach:

#### 1. Primary Check (Payment Pipeline Only)
- **If 100% Sales-owned payments** → "Sales Team"
- **If 100% CS-owned payments** → "CS Team"
- **If mixed** → proceed to fallback

#### 2. Fallback (All Pipelines, All Time)
Scans all deals across Sales + CS + Payment pipelines:
- **If all historical deals are Sales** → "Sales Team"
- **If all historical deals are CS** → "CS Team"
- **If mixed** → proceed to tie-breaker

#### 3. Tie-Breaker (Current FY Revenue Only)
Looks at current fiscal year revenue deals:
- **If purely Sales** → "Sales Team"
- **If purely CS** → "CS Team"
- **If mixed** → "CS and Sales (Transferred this FY)"

#### 4. Final Defaults
- **No current FY revenue AND mixed history** → "CS & Sales"
- **No Sales/CS ownership flags** → "C-Suite"

### Revenue Calculations

#### Realized Revenue (This FY)
- Sum of actual payments with Close Date within current fiscal year
- Close Date ≤ Report Date
- **Excludes forecasted amounts**

#### Forecasted Revenue
**Only for Recurring Revenue companies**

**Sales Team & Summary Report Rules:**
- Forecast only if `paymentCount < 12`
- Start forecasting from month AFTER Report Date
- Maximum 12 total months (including already paid)
- Amount = `lastPaymentAmount`
- **Do NOT forecast if:**
  - Revenue in Report Date month = 0, OR
  - Revenue in month immediately before forecast month = 0

**CS Team Report Rules:**
- **No 12-month limit**
- Forecast until EARLIER of:
  - Report End Date (E9), OR
  - June 30 of current fiscal year
- All other rules identical to Sales/Summary

**Example:**
```
Company: Recurring Revenue
Last Payment: $1,000 in November 2024
Report Date: 12/31/2024
Payment Count: 8

Forecast:
Jan 2025: $1,000
Feb 2025: $1,000
Mar 2025: $1,000
Apr 2025: $1,000 (stops at 12 total months for Sales report)
```

---

## 📊 Report Specifications

### Summary Report

**Sheet Name:** `Summary Report`

**Purpose:** Complete view of all companies without filtering

**Columns:**
1. Company Name
2. Revenue Trend (Sparkline)
3. Revenue Type (Recurring/R-OTP/OTP)
4. Client Age (New/Old/Future)
5. First Revenue Fiscal Year
6. POC Team
7. Monthly columns (YYYY-MMM format from Report Start to Report End)
8. Realized Revenue (This FY)
9. Total Revenue

**Features:**
- Auto-sort by First Revenue FY, then alphabetically
- Conditional formatting:
  - Light yellow background for rows with current FY activity
  - Grey font for non-current FY months
  - Green font for month-over-month increase (>10% or >$50)
  - Red font for month-over-month decrease (>10% or >$50)
- Frozen first row and first 3 columns
- Auto-filter on headers
- Subtotal row using `SUBTOTAL(9, ...)` (filter-safe)

### Sales Team Report

**Sheet Name:** `Sales Team Report`

**Purpose:** Filtered view for Sales team with growth analysis

**Additional Filtering:**

#### Filter A: Company Exclusion List
- **Source:** `Company_Stop_list` sheet, Column A
- **Logic:** EXCLUDE companies in this list

#### Filter B: Sales Team Membership
- **Source:** `TEMP_DATA` sheet, Column Z "Company has Sales Team"
- **Logic:** Include only if TRUE

#### Filter C: Revenue This FY (Optional)
- **Source:** `Config_sheet` E12
- **Logic:** If E12 = TRUE, include only companies with "Has Revenue this FY" = TRUE

#### Filter D: Aggregation Level (Optional)
- **Source:** `Config_sheet` E13
- **Options:**
  - "Deal Level" → Aggregate only deals where "Is Sales Team Member" = TRUE
  - "Company Level" → Standard aggregation (default)

#### Filter E: New Only (Optional)
- **Source:** `Config_sheet` E14
- **Logic:** If E14 = "New only", include only "Client Age" = "New"

#### Filter F: OLD OTP Payment Threshold
- **Applies to:** Client Age = "OLD" AND Deal Type = "OTP"
- **Source:** `Config_sheet` E15 (default: 5)
- **Logic:** INCLUDE if payments in one FY ≤ threshold

#### Filter G: OLD Recurring Growth Check
- **Applies to:** Client Age = "OLD" AND Deal Type = "Recurring" or "R-OTP"
- **Source:** `Config_sheet` E16 (checkbox)
- **Logic:** If E16 = TRUE, INCLUDE only if:
  - ≥2 paid months in current FY, AND
  - At least ONE month shows increase (>10% or >$50)
  - EXCLUDE if ≤1 paid month in current FY

#### Filter H: Transferred Account Window
- **Applies to:** Deal Type = "Recurring" or "R-OTP", AND "Is CS Team Member" = TRUE
- **Source:** `Config_sheet` E17 (default: 12 months)
- **Logic:** Include deals where Close Date is within X months after First Revenue Month

**Forecasting:**
- Same rules as Summary Report
- Maximum 12 months total (including actual payments)
- Uses `lastPaymentAmount`

### CS Team Report

**Sheet Name:** `CS Team Report`

**Purpose:** CS-focused view with extended forecasting

**Filtering:**
- **CS Team Membership:** Include only companies with "Is CS Team Member" = TRUE
- **Exclude:** Configurable list (e.g., "Sales Team", "C-Suite")

**Forecasting:**
- **No 12-month limit** (key difference from Sales)
- Forecast until EARLIER of:
  - Report End Date (E9)
  - June 30 of current fiscal year
- Carry-forward logic:
  - Use Report Month amount if > 0
  - Else use prior month if > 0
  - Else no forecast
- Forecasted months shown in grey font

**Columns:**
1. Company Name
2. Revenue Type
3. Client Age
4. First Revenue Fiscal Year
5. POC Team
6. Monthly columns (YYYY-MMM)
7. Realized Revenue (This FY)
8. Forecasted Revenue (This FY only)
9. Total Report Period Revenue

---

## 🔧 Configuration Reference

### CONFIG Object

Located in the `Config` file:

```javascript
const CONFIG = {
  // HubSpot Authentication
  HUBSPOT_TOKEN: 'your-token-here',
  DEFAULT_START: '2023-01-01',

  // Pipeline Definitions
  PIPELINE_MAP: {
    '70710959': 'Sales Pipeline',
    '679755281': 'CS Pipeline',
    'default': 'Payment Pipeline'
  },

  // Sheet Names
  TEMP_SHEET_NAME: 'TEMP_DATA',
  CONFIG_SHEET_NAME: 'Config_sheet',
  SALES_TEAM_SHEET: 'Sales_team_Members',
  CS_TEAM_SHEET: 'CS_team_Members',
  PAYMENT_SHEET_NAME: 'Payment Pipeline',

  // Fiscal Year Settings
  FISCAL_YEAR_START_MONTH: 6,  // July = 6 (0-indexed)

  // Performance Settings
  BATCH_SIZE: 2500,
  COOLDOWN_MS: 2000,

  // Column Indices
  PIPELINE_HEADER_ROW: 2,
  TEAM_HEADER_ROW: 1,
  REPORT_DATE_CELL: 'E6'
};
```

### Stage Mapping

```javascript
const STAGE_MAP = {
  '136833349': 'Closed Won (Sales Pipeline)',
  'closedwon': 'Revenue (Payment Pipeline)',
  '996632637': 'Closed Won (Payment Pipeline)'
};
```

### Classification Thresholds

Default values (configurable in code):

```javascript
recurringMonthsThresholdFY: 8     // Months in one FY to be Recurring
rOtpMonthsThresholdTotal: 5       // Total months to be R-OTP
revenueIncreaseThreshold: 0.10    // 10% increase threshold
revenueIncreaseDollarThreshold: 50 // $50 increase threshold
```

---

## 🎨 Formatting Rules

### Color Coding

**Background Colors:**
- Light Yellow (#FFFFE0) - Rows with current FY activity

**Font Colors:**
- Green (#008000) - Month-over-month increase (>10% or >$50, current FY only)
- Red (#FF0000) - Month-over-month decrease (>10% or >$50, current FY only)
- Light Grey (#D3D3D3) - Months outside current fiscal year
- Grey (Forecast) - Forecasted months in CS Team Report

### Number Formatting

**Currency:**
```
Format: $###,###
Display: $1,234 or $1,234.00
Zero Value: "-" (dash, center-aligned)
```

**Headers:**
```
Background: Dark Grey (#444444)
Font Color: White
Font Weight: Bold
```

### Layout

**Frozen Areas:**
- Row 1 (Headers) - Always frozen
- Columns A:C (Company Name + first columns) - Configurable

**Auto-Filter:**
- Applied to entire data range
- Allows dynamic filtering without breaking formulas

**Auto-Resize:**
- All columns auto-resize after data population
- Ensures readability without manual adjustment

---

## 🔍 Troubleshooting

### Common Issues

#### 1. "Sheet not found" Error
**Cause:** Required sheets don't exist  
**Solution:** Create these sheets manually:
- Config_sheet
- Sales_team_Members
- CS_team_Members
- Company_Stop_list

#### 2. No Data Imported
**Cause:** Invalid HubSpot token or incorrect stage IDs  
**Solution:** 
- Verify HUBSPOT_TOKEN in Config file
- Check STAGE_IDS match your HubSpot deal stages
- Review Apps Script execution logs

#### 3. Duplicate Deals
**Cause:** Multiple syncs without deduplication  
**Solution:** The Temp Sheet process automatically deduplicates by Deal ID (most recent Last Modified Date wins)

#### 4. Missing Company Names
**Cause:** Company associations not synced  
**Solution:** Run "2. Sync Missing Companies" from menu

#### 5. Forecast Not Showing
**Cause:** Company not classified as Recurring Revenue  
**Solution:** 
- Check Revenue Type classification
- Verify payment count ≥ 8 months in one FY OR deal name contains keywords
- Check that paymentCount < 12

#### 6. API Rate Limit Errors
**Cause:** Too many requests to HubSpot API  
**Solution:** 
- Increase COOLDOWN_MS value (default: 2000ms)
- Reduce BATCH_SIZE (default: 100)
- Run syncs less frequently

---

## 🚦 Performance Optimization

### Batch Processing

**HubSpot API Calls:**
- 100 deals per batch (configurable via BATCH_SIZE)
- 300ms sleep between batches (configurable via COOLDOWN_MS)
- Pagination handling for large datasets

**Sheet Writing:**
- Write in chunks of 2,500 rows (CONFIG.BATCH_SIZE)
- Prevents timeout on large datasets
- 2-second cooldown between chunks

### Caching Strategy

**TEMP_DATA Sheet:**
- Hidden sheet stores pre-computed classifications
- Eliminates redundant calculations across reports
- All reports use same cached data
- Only regenerate when source data changes

**Performance Gains:**
```
Without Cache: 3-5 minutes per report
With Cache: 15-30 seconds per report
```

### Best Practices

1. **Run Temp Sheet generation first** before any reports
2. **Schedule syncs during off-hours** to minimize UI lag
3. **Use filters in reports** instead of re-syncing data
4. **Archive old data** to separate sheets if exceeding 10,000 rows
5. **Monitor API usage** in HubSpot to stay within limits

---

## 📈 Use Cases

### Perfect For

✅ **Sales Teams**
- Track deal performance across pipelines
- Monitor new business vs expansions
- Commission calculations
- Pipeline forecasting

✅ **Customer Success Teams**
- Monitor customer health
- Renewal forecasting
- Account transition tracking
- Churn risk analysis

✅ **Finance Teams**
- Revenue recognition
- Cash flow projections
- Budget vs actual analysis
- Recurring revenue metrics

✅ **Operations Teams**
- Regular CRM data exports
- Cross-functional reporting
- Data validation
- Ad-hoc analysis

✅ **Non-Technical Users**
- No SQL or coding required
- Intuitive custom menu
- Automatic formatting
- Filter-friendly outputs

---

## 📝 Version History

### Current Version: 2.0

**Features:**
- Multi-pipeline sync (Sales, CS, Payment)
- Automated classification (Revenue Type, Client Age, POC Team)
- Three report types (Summary, Sales Team, CS Team)
- Advanced forecasting logic
- TLD-based company name mapping
- Configurable filtering rules
- Batch processing with rate limiting
- Hidden cache for performance
- Custom menu interface

**Recent Updates:**
- Added CS Team Report with extended forecasting
- Implemented hierarchical POC team classification
- Enhanced Sales Team filtering with growth logic
- Added TLD mapping for company name enrichment
- Improved deduplication by Last Modified Date
- Added deal-level vs company-level aggregation
- Implemented configurable thresholds
- Enhanced error handling and logging

---

## 🛠️ Development & Customization

### Adding New Pipelines

To add a new pipeline:

1. **Update CONFIG in Config file:**
```javascript
PIPELINE_MAP: {
  '70710959': 'Sales Pipeline',
  '679755281': 'CS Pipeline',
  'your-pipeline-id': 'Your Pipeline Name',
  'default': 'Payment Pipeline'
}
```

2. **Update SOURCE_SHEETS array:**
```javascript
SOURCE_SHEETS: [
  'Sales Pipeline',
  'CS Pipeline',
  'Your Pipeline Name',
  'Payment Pipeline'
]
```

3. **Add stage IDs to STAGE_IDS if needed:**
```javascript
const STAGE_IDS = [
  '136833349',
  'closedwon',
  '996632637',
  'your-stage-id'
];
```

### Adding New Team Lists

To add a new team classification:

1. **Create a new sheet** (e.g., "Marketing_team_Members")
2. **Add Owner IDs** in Column A (with header in Row 1)
3. **Update Temp Sheet.gs** to read the new team list
4. **Add classification logic** in the main processing function

### Customizing Reports

#### Change Column Order

Edit the `EXTRA_COLUMNS` array in report configs:

```javascript
EXTRA_COLUMNS: [
  { header: "Revenue Trend", isSparkline: true },
  { header: "Revenue Type", index: 15 },
  { header: "Client Age", index: 16 },
  { header: "Your New Column", index: 25 }  // Add here
]
```

#### Modify Conditional Formatting

Edit the formatting sections in report functions:

```javascript
// Change threshold for green highlighting
if (curr > prev * 1.15 || curr > prev + 100) {  // Was 1.1 and 50
  return "green";
}
```

#### Add New Filters

In Sales Team Report or CS Team Report, add filter logic:

```javascript
// Example: Filter by deal amount
if (dealAmount < 1000) {
  return false;  // Exclude deals under $1,000
}
```

---

## 🔐 Security & Privacy

### Data Handling

- **API Token:** Stored in script properties (not visible in sheet)
- **Data Scope:** Only reads deals with specific stage IDs
- **Permissions:** Requires HubSpot read-only access
- **Storage:** All data stored in your Google Sheet (you own the data)

### Best Practices

1. **Use Private Apps** in HubSpot (not personal tokens)
2. **Limit token scopes** to read-only on required objects
3. **Restrict sheet access** using Google Sheets permissions
4. **Rotate tokens** periodically (every 90 days recommended)
5. **Audit logs** regularly in HubSpot for API usage
6. **Delete historical data** if no longer needed

---

## 📞 Support & Maintenance

### Debugging

Enable detailed logging:

```javascript
// In any function, add:
Logger.log('Debug info:', variable);

// View logs:
// Apps Script Editor → Executions → View logs
```

Use the built-in debugger:

```javascript
function debugConfig() {
  console.log('--- START CONFIG CHECK ---');
  console.log('HubSpot Token exists:', !!CONFIG.HUBSPOT_TOKEN);
  // ... additional checks
}
```

### Monitoring

**Key Metrics to Track:**
- Total deals synced per run
- API calls per sync
- Execution time per report
- Error count in logs
- HubSpot API usage (via HubSpot dashboard)

**Set Up Alerts:**
- Create time-driven triggers for automatic syncs
- Add email notifications on errors:

```javascript
function sendErrorEmail(error) {
  MailApp.sendEmail({
    to: 'your-email@example.com',
    subject: 'HubSpot Sync Error',
    body: error.message
  });
}
```

---

## 🤝 Contributing

This is an internal tool, but improvements are welcome!

### Code Style

- Use descriptive variable names
- Add comments for complex logic
- Follow existing naming conventions
- Test thoroughly before deploying

### Submitting Changes

1. Test in a copy of the spreadsheet
2. Document all changes in code comments
3. Update this README if adding features
4. Share with team for review

---

## 📄 License

Internal use only. Not for public distribution.

---

## 📚 Additional Resources

### HubSpot API Documentation
- [HubSpot API Overview](https://developers.hubspot.com/docs/api/overview)
- [Deals API](https://developers.hubspot.com/docs/api/crm/deals)
- [Associations API](https://developers.hubspot.com/docs/api/crm/associations)
- [Search API](https://developers.hubspot.com/docs/api/crm/search)

### Google Apps Script Documentation
- [Apps Script Overview](https://developers.google.com/apps-script)
- [Spreadsheet Service](https://developers.google.com/apps-script/reference/spreadsheet)
- [URL Fetch Service](https://developers.google.com/apps-script/reference/url-fetch)
- [Utilities Service](https://developers.google.com/apps-script/reference/utilities)

### Related Tools
- [HubSpot Private Apps](https://developers.hubspot.com/docs/api/private-apps)
- [Google Sheets Functions Reference](https://support.google.com/docs/table/25273)

---

**Last Updated:** February 16, 2025  
**Maintained By:** Internal Development Team  
**Version:** 2.0
