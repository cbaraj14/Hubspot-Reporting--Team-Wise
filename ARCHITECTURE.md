# System Architecture

Visual guide to understanding the HubSpot Deals to Google Sheets Importer architecture.

---

## 📐 High-Level Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                         HubSpot CRM                             │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐         │
│  │    Sales     │  │      CS      │  │   Payment    │         │
│  │   Pipeline   │  │   Pipeline   │  │   Pipeline   │         │
│  │  (70710959)  │  │ (679755281)  │  │   (default)  │         │
│  └──────────────┘  └──────────────┘  └──────────────┘         │
└────────────────────────────┬────────────────────────────────────┘
                             │
                    ┌────────▼────────┐
                    │   HubSpot API   │
                    │    (REST v3)    │
                    └────────┬────────┘
                             │
              ╔══════════════▼══════════════╗
              ║     Google Apps Script      ║
              ║    hubspot_sync.gs          ║
              ║  - Fetch deals (batch 100)  ║
              ║  - Enrich associations      ║
              ║  - Deduplicate by ID        ║
              ╚══════════════╤══════════════╝
                             │
        ┌────────────────────┼────────────────────┐
        │                    │                    │
   ┌────▼─────┐      ┌──────▼──────┐      ┌─────▼─────┐
   │  Sales   │      │     CS      │      │  Payment  │
   │ Pipeline │      │  Pipeline   │      │  Pipeline │
   │  Sheet   │      │   Sheet     │      │   Sheet   │
   └────┬─────┘      └──────┬──────┘      └─────┬─────┘
        └────────────────────┼────────────────────┘
                             │
              ╔══════════════▼══════════════╗
              ║     Temp Sheet.gs           ║
              ║  - Classify Revenue Type    ║
              ║  - Classify Client Age      ║
              ║  - Calculate Fiscal Years   ║
              ║  - Assign Teams             ║
              ╚══════════════╤══════════════╝
                             │
                      ┌──────▼──────┐
                      │  TEMP_DATA  │
                      │   (Hidden)  │
                      │  25+ cols   │
                      └──────┬──────┘
                             │
        ┌────────────────────┼────────────────────┐
        │                    │                    │
   ┌────▼─────┐      ┌──────▼──────┐      ┌─────▼─────┐
   │ Summary  │      │    Sales    │      │    CS     │
   │  Report  │      │Team Report  │      │Team Report│
   │  (All)   │      │ (Filtered)  │      │(Forecast) │
   └──────────┘      └─────────────┘      └───────────┘
```

---

## 🔄 Data Flow Sequence

### Phase 1: HubSpot Sync

```
1. User triggers: HubSpot Sync → 1. Sync New Deals

2. Script reads: Last sync timestamp from Sheet A1

3. API Request:
   POST https://api.hubapi.com/crm/v3/objects/deals/search
   {
     filterGroups: [
       { filters: [
           { propertyName: 'hs_lastmodifieddate', operator: 'GTE', value: <timestamp> },
           { propertyName: 'dealstage', operator: 'IN', values: ['136833349', 'closedwon', '996632637'] }
       ]}
     ],
     properties: ['dealname', 'amount', 'closedate', ...],
     limit: 100,
     after: <pagination_token>
   }

4. Response handling:
   - Parse JSON response
   - Group by pipeline ID
   - Batch write to respective sheets (100 rows at a time)
   - Sleep 300ms between batches

5. Repeat step 3-4 until no more pages

6. Update "Last Updated" timestamp in A1
```

### Phase 2: Association Enrichment

```
1. User triggers: HubSpot Sync → 2. Sync Missing Companies

2. Script scans pipeline sheets for empty Company ID (col G)

3. For each batch of 100 deal IDs:
   POST https://api.hubapi.com/crm/v3/associations/deals/companies/batch/read
   { inputs: [ {id: '123'}, {id: '456'}, ... ] }

4. Fill Company ID column with results

5. If no company found:
   - Check if Contact Name exists → use "Individual-(Name)"
   - Else use Contact Email as fallback

6. Repeat for contacts (similar flow)
```

### Phase 3: ID to Name Conversion

```
1. User triggers: HubSpot Sync → 4. Update ID to Names

2. Script collects all unique IDs:
   - Company IDs (col G)
   - Contact IDs (col H)
   - Owner IDs (col F)

3. For each ID type, batch API call:
   GET https://api.hubapi.com/crm/v3/objects/{type}/{id}?properties=name,email

4. Build lookup maps:
   companyIdToName = { '123': 'Acme Corp', ... }
   contactIdToName = { '456': 'John Doe', ... }
   ownerIdToName = { '789': 'Jane Smith', ... }

5. Update pipeline sheets with names (cols L, M, N, O)
```

### Phase 4: Classification & Caching

```
1. User triggers: HubSpot Sync → (runs before reports)

2. tempUpdateTempSheet() executes:

   a) Read all pipeline sheets
   
   b) Deduplicate by Deal ID (keep most recent Last Modified)
   
   c) For each company, calculate:
      - First payment date
      - Unique payment months (per FY and total)
      - Team attribution flags
   
   d) Classify each deal:
      Revenue Type = classify_revenue_type(deal_name, month_count)
      Client Age = classify_client_age(first_payment, report_date, FY)
      POC Team = assign_poc_team(owner_flags)
   
   e) Write to TEMP_DATA (hidden sheet):
      - All original columns
      - Plus 10+ computed columns
   
   f) Hide TEMP_DATA sheet to avoid confusion

3. All reports now read from TEMP_DATA (fast!)
```

### Phase 5: Report Generation

```
SUMMARY REPORT:
1. Read TEMP_DATA (all rows)
2. Filter: Pipeline = "Payment Pipeline" AND Date in range
3. Pivot by Company Name × Month
4. Calculate POC Team (hierarchical logic)
5. Generate monthly columns (YYYY-MMM)
6. Add sparklines and conditional formatting
7. Write to "Summary Report" sheet

SALES TEAM REPORT:
1. Read TEMP_DATA
2. Apply mandatory filters:
   - NOT IN Company_Stop_list
   - Company has Sales Team = TRUE
3. Apply optional filters (based on Config_sheet):
   - Revenue This FY (E12)
   - Aggregation Level (E13)
   - New Clients Only (E14)
   - Payment Thresholds (E15, E16, E17)
4. Same pivot/formatting as Summary
5. Write to "Sales Team Report" sheet

CS TEAM REPORT:
1. Read "Summary Report" (already aggregated)
2. Filter: POC Team NOT IN ["Sales Team", "C-Suite"]
3. For Recurring companies:
   - Forecast from Report Month → Report End
   - Use last payment amount
   - No 12-month limit (extends to FY end)
4. Calculate:
   - Realized Revenue (actual only)
   - Forecasted Revenue (actual + forecast)
   - Total Revenue (all months)
5. Grey out forecast months
6. Write to "Summary_CS_Team" sheet
```

---

## 🗂️ Data Model

### Pipeline Sheets (Sales/CS/Payment)

**Structure:** 15 columns × N rows (raw HubSpot data)

| Index | Column | Type | Source | Example |
|-------|--------|------|--------|---------|
| 0 | Deal Name | Text | HubSpot | "Q4 Renewal - Acme Corp" |
| 1 | Amount | Number | HubSpot | 5000 |
| 2 | Date Created | Date | HubSpot | 2024-01-15 |
| 3 | Stage Name | Text | Mapped | "Closed Won (Sales)" |
| 4 | Pipeline Name | Text | Mapped | "Sales Pipeline" |
| 5 | Owner ID | Text | HubSpot | "12345678" |
| 6 | Associated Company ID | Text | HubSpot | "98765432" |
| 7 | Associated Contact ID | Text | HubSpot | "11223344" |
| 8 | Deal ID | Text | HubSpot | "55667788" |
| 9 | Close Date | Date | HubSpot | 2024-03-20 |
| 10 | Last Modified Date | Date | HubSpot | 2024-03-21 |
| 11 | Company Name | Text | Enriched | "Acme Corporation" |
| 12 | Contact Name | Text | Enriched | "John Doe" |
| 13 | Owner Name | Text | Enriched | "Jane Smith" |
| 14 | Contact Email | Text | Enriched | "john@acme.com" |

**Keys:**
- Primary: Deal ID (col 8)
- Foreign: Company ID (col 6), Contact ID (col 7), Owner ID (col 5)

**Constraints:**
- Deal ID must be unique
- Close Date required for revenue calculations
- Pipeline Name determines processing rules

### TEMP_DATA Sheet (Hidden Cache)

**Structure:** 25+ columns × N rows (enriched data)

| Index | Column | Type | Computed | Example |
|-------|--------|------|----------|---------|
| 0-14 | Original columns | Various | From pipelines | (see above) |
| 15 | Deal Type | Text | ✅ | "Recurring" |
| 16 | Client Age | Text | ✅ | "Old" |
| 17 | First Revenue Month | Text | ✅ | "2023-Jan" |
| 18 | First Revenue Fiscal Year | Text | ✅ | "FY 22/23 Q3" |
| 19 | Realized Revenue (this FY) | Number | ✅ | 15000 |
| 20 | Is Sales Team Member | Boolean | ✅ | TRUE |
| 21 | Has Revenue this FY | Boolean | ✅ | TRUE |
| 22 | Has Sales Team Revenue This FY | Boolean | ✅ | TRUE |
| 23 | Is CS Team Member | Boolean | ✅ | FALSE |
| 24 | Has CS Team Revenue This FY | Boolean | ✅ | FALSE |
| 25 | Company has Sales Team | Boolean | ✅ | TRUE |
| 26 | Has Revenue this FY | Boolean | ✅ | TRUE |
| 27 | Any Deals this FY | Boolean | ✅ | TRUE |

**Purpose:**
- Pre-computed classifications (run once, use many times)
- Single source of truth for all reports
- Hidden to avoid user confusion
- Regenerated when source data changes

### Report Sheets (Summary/Sales Team/CS Team)

**Structure:** Dynamic columns × N rows (aggregated data)

**Common columns:**
1. Company Name
2. Revenue Trend (Sparkline)
3. Revenue Type
4. Client Age
5. First Revenue FY
6. POC Team
7-N. Monthly columns (YYYY-MMM from start to end date)
N+1. Realized Revenue (This FY)
N+2. Forecasted Revenue (CS only)
N+3. Total Revenue

**Last row:** SUBTOTAL formulas (filter-safe)

---

## 🔐 Security Architecture

### Authentication Flow

```
┌──────────────┐
│     User     │
│  Opens Sheet │
└──────┬───────┘
       │
       ▼
┌──────────────────┐
│   onOpen()       │
│ Creates Menu     │
└──────┬───────────┘
       │
       ▼
┌──────────────────────────┐
│ User clicks menu item    │
└──────┬───────────────────┘
       │
       ▼ (First time only)
┌──────────────────────────┐
│ OAuth Authorization      │
│ - Request permissions    │
│ - User approves          │
│ - Token stored           │
└──────┬───────────────────┘
       │
       ▼
┌──────────────────────────┐
│ Function executes        │
│ - Reads CONFIG token     │
│ - Calls HubSpot API      │
│ - Updates sheets         │
└──────────────────────────┘
```

### Token Storage

```javascript
// ✅ SECURE - Token in script properties (encrypted)
const HUBSPOT_TOKEN = PropertiesService
  .getScriptProperties()
  .getProperty('HUBSPOT_TOKEN');

// ⚠️ LESS SECURE - Token in code (visible in editor)
const HUBSPOT_TOKEN = 'pat-na1-xxxxxx';  // Current implementation
```

**Recommendation:** Migrate to PropertiesService for production

### API Security

```
┌─────────────┐
│   Script    │
└─────┬───────┘
      │
      ▼
┌─────────────────────────┐
│  urlFetchWithRetry()    │
│  - Adds Bearer token    │
│  - Rate limiting        │
│  - Exponential backoff  │
│  - Error handling       │
└─────┬───────────────────┘
      │
      ▼
┌─────────────────────────┐
│   HubSpot API           │
│  - Validates token      │
│  - Checks scopes        │
│  - Rate limits applied  │
│  - Returns data/error   │
└─────────────────────────┘
```

---

## ⚙️ Configuration Architecture

### Config Hierarchy

```
┌─────────────────────────────────────────┐
│         CONFIG (Global Object)          │
│  - Master settings                      │
│  - Used by all modules                  │
│  - Defined once in Config file          │
└─────────────┬───────────────────────────┘
              │
    ┌─────────┼─────────┬─────────────┐
    │         │         │             │
┌───▼──┐  ┌──▼───┐  ┌──▼───┐  ┌─────▼──────┐
│SUMMARY│ │SALES │ │  CS  │  │   temp     │
│CONFIG │ │CONFIG│ │CONFIG│  │  CONFIG    │
└───────┘ └──────┘ └──────┘  └────────────┘
  │          │        │            │
  │          │        │            │
  ▼          ▼        ▼            ▼
Report    Report   Report    Processing
Logic     Logic    Logic      Logic
```

### Config Sheets Structure

```
Config_sheet
├── E6:  Report Date          → Anchor for current month
├── E8:  Report Start Date    → Begin date range
├── E9:  Report End Date      → End date range
├── E12: Sales Revenue Filter → TRUE/FALSE
├── E13: Sales Aggregation    → Deal/Company Level
├── E14: Sales Age Filter     → New only/All
├── E15: Payment Threshold    → Number (5)
├── E16: Growth Check         → TRUE/FALSE
└── E17: Transfer Window      → Number (12)

Sales_team_Members
├── A1: Owner ID (header)
└── A2+: Owner IDs

CS_team_Members
├── A1: Owner ID (header)
└── A2+: Owner IDs

Company_Stop_list
├── A1: Company Name (header)
└── A2+: Company names to exclude
```

---

## 🚀 Performance Architecture

### Optimization Strategies

**1. Batch Processing**
```
Single API Call (slow)        Batch API Call (fast)
─────────────────────        ────────────────────────
Call 1: Get deal 1           Call 1: Get deals 1-100
Call 2: Get deal 2           Call 2: Get deals 101-200
...                          ...
Call 1000: Get deal 1000     Call 10: Get deals 901-1000

Time: ~1000 × 200ms = 200s   Time: 10 × 500ms = 5s
```

**2. Caching Strategy**
```
Without TEMP_DATA             With TEMP_DATA
─────────────────            ───────────────
Summary: Process 10,000      Generate TEMP_DATA once:
  rows, classify, filter      - Process 10,000 rows
  → 3 minutes                 - Classify all
                              - Cache results
Sales: Process 10,000          → 3 minutes (one-time)
  rows, classify, filter
  → 3 minutes               Summary: Read cache → 30s
                            Sales: Read cache → 30s
CS: Process 10,000          CS: Read cache → 30s
  rows, classify, filter
  → 3 minutes               Total: 4 minutes
                            (vs 9 minutes without cache)
Total: 9 minutes
```

**3. Sheet Writing Optimization**
```
Row-by-Row (slow)             Batch Write (fast)
─────────────────            ──────────────────
for (row in data) {          const batchSize = 2500;
  sheet.appendRow(row);      for (i = 0; i < rows; i += batchSize) {
}                              const chunk = rows.slice(i, i + batchSize);
                               sheet.getRange(...).setValues(chunk);
Time: N × 50ms                }
For 10k rows: 500s           Time: (N/batchSize) × 500ms
                             For 10k rows: 2s
```

---

## 🧩 Module Dependencies

```
                  ┌───────────┐
                  │  Menu.gs  │
                  └─────┬─────┘
                        │
        ┌───────────────┼───────────────┐
        │               │               │
   ┌────▼────┐    ┌────▼────┐    ┌────▼────┐
   │ HubSpot │    │  Temp   │    │ Report  │
   │ Sync    │    │  Sheet  │    │ Modules │
   └────┬────┘    └────┬────┘    └────┬────┘
        │              │              │
        └──────────────┼──────────────┘
                       │
                  ┌────▼────┐
                  │ Config  │
                  └─────────┘
```

**Dependencies:**
- All modules depend on `Config`
- Reports depend on `Temp Sheet` output
- `Menu.gs` orchestrates all modules
- No circular dependencies

---

## 📊 Scalability Considerations

### Current Limits

| Metric | Limit | Reason |
|--------|-------|--------|
| Max rows per sheet | ~100,000 | Google Sheets limit |
| Execution time | 6 minutes | Apps Script limit |
| API calls/day | ~10,000 | HubSpot rate limit |
| Memory | 100 MB | Apps Script limit |

### Scaling Strategies

**For Large Datasets (50k+ deals):**
1. Increase batch size to 5,000
2. Archive old data to separate sheets
3. Narrow date ranges
4. Run syncs overnight (time-based triggers)

**For High Frequency (hourly updates):**
1. Reduce sync scope (only new deals)
2. Use webhooks instead of polling (future)
3. Implement incremental updates
4. Cache more aggressively

**For Many Users:**
1. Create separate copies per team
2. Centralize data sync, distribute reports
3. Use Google Sheets sharing (don't duplicate)
4. Implement row-level permissions (manual)

---

## 🔮 Future Enhancements

### Planned Architecture Changes

**Phase 1: Performance**
- Move to BigQuery for storage (10M+ row support)
- Implement server-side caching
- Add data archiving automation

**Phase 2: Real-time**
- Implement HubSpot webhooks
- Add real-time deal updates
- Create live dashboards

**Phase 3: Enterprise**
- Multi-tenant support
- Role-based access control
- Audit logging
- API endpoints for external tools

---

*Last Updated: February 16, 2025*
*Version: 2.0.0*
