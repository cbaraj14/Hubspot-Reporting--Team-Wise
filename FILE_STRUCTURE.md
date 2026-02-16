# File Structure

Complete listing of all files in the HubSpot Deals to Google Sheets Importer repository.

---

## 📄 Documentation Files

### README.md (29 KB)
**Purpose:** Comprehensive documentation for the entire system  
**Contents:**
- Feature overview and key capabilities
- Architecture diagrams and data flow
- Complete setup instructions
- Business logic specifications
- Configuration reference
- Report specifications
- Troubleshooting guide
- Performance optimization tips
- Security best practices

**Audience:** All users - developers and non-technical users

---

### QUICKSTART.md (7.4 KB)
**Purpose:** Fast-track setup guide for new users  
**Contents:**
- 5-step setup process (15 minutes)
- Prerequisites checklist
- HubSpot API token setup
- Google Sheets configuration
- First sync walkthrough
- Verification steps
- Quick troubleshooting

**Audience:** New users who want to get started quickly

---

### ARCHITECTURE.md (20 KB)
**Purpose:** Visual and technical architecture documentation  
**Contents:**
- High-level architecture diagrams
- Data flow sequence diagrams
- Data model specifications
- Security architecture
- Configuration hierarchy
- Performance optimization patterns
- Module dependencies
- Scalability considerations

**Audience:** Developers and technical stakeholders

---

### CHANGELOG.md (6.4 KB)
**Purpose:** Version history and change tracking  
**Contents:**
- Version 2.0.0 documentation update
- Historical features from v1.5.0
- Future roadmap
- Breaking changes log
- Maintenance notes

**Audience:** All users tracking version changes

---

### CONTRIBUTING.md (13 KB)
**Purpose:** Guide for team members making changes  
**Contents:**
- Development setup instructions
- Code style guide
- Testing checklist
- Common tasks (adding filters, columns, reports)
- Debugging tips
- Submission process
- Common pitfalls to avoid

**Audience:** Developers and contributors

---

### LICENSE
**Purpose:** Internal use license and compliance  
**Contents:**
- Internal use restrictions
- Data security requirements
- Confidentiality terms
- Compliance guidelines
- Third-party service terms

**Audience:** Legal, compliance, and all users

---

### FILE_STRUCTURE.md (This file)
**Purpose:** Complete file listing and descriptions  
**Contents:**
- All documentation files
- All code files
- File sizes and purposes
- Quick reference guide

**Audience:** All users navigating the codebase

---

## 💻 Code Files (Google Apps Script)

### Config (6.4 KB)
**Purpose:** Global configuration object - single source of truth  
**Key sections:**
- HubSpot authentication token
- Pipeline ID mappings
- Sheet names (input, cache, output)
- Fiscal year settings (July 1 start)
- Performance tuning (batch size, delays)
- Column structure definitions
- Stage ID mappings

**Key variables:**
```javascript
CONFIG = {
  HUBSPOT_TOKEN: 'your-token-here',
  PIPELINE_MAP: { ... },
  SOURCE_SHEETS: [ ... ],
  FISCAL_YEAR_START_MONTH: 6,
  BATCH_SIZE: 2500,
  COOLDOWN_MS: 2000
}
```

**Modified by:** Developers during setup
**Read by:** All modules

---

### hubspot_sync.gs (16 KB)
**Purpose:** HubSpot API integration and data import  
**Key functions:**
- `syncHubSpotDeals()` - Import new/updated deals
- `syncMissingCompanies()` - Enrich company associations
- `syncMissingContacts()` - Enrich contact associations
- `updateLastModifiedForExistingDeals()` - Refresh timestamps
- `urlFetchWithRetry()` - API call with retry logic

**Features:**
- Batch processing (100 deals per call)
- Pagination handling
- Rate limiting (300ms between calls)
- Exponential backoff on errors
- Automatic deduplication

**Accessed via:** Menu items 1, 2, 3

---

### Temp Sheet.gs (14 KB)
**Purpose:** Data classification and enrichment engine  
**Key functions:**
- `tempUpdateTempSheet()` - Main processing function
- `tempClassifyDealType()` - Revenue Type classification
- `tempClassifyAge()` - Client Age classification
- `tempCalculateCompanyStats()` - Aggregate company metrics
- `tempGetFiscalYear()` - Fiscal year calculations

**Operations:**
1. Deduplicate by Deal ID
2. Classify Revenue Type (Recurring/R-OTP/OTP)
3. Classify Client Age (New/Old/Future)
4. Calculate fiscal years
5. Assign team attribution
6. Write to hidden TEMP_DATA sheet

**Output:** 25+ column cache sheet
**Accessed via:** Menu item 4 (before reports)

---

### Summary reprot.gs (12 KB)
**Purpose:** Generate comprehensive summary report (all companies)  
**Key function:** `generateSummaryReport()`

**Features:**
- No filtering (all companies included)
- POC Team hierarchical classification
- Monthly revenue pivot
- Sparkline visualizations
- Realized vs Total revenue
- Conditional formatting (colors, fonts)

**Output sheet:** "Summary Report"
**Accessed via:** Menu item 5

---

### Sales Team Report.gs (18 KB)
**Purpose:** Generate filtered Sales team report  
**Key function:** `salesGenerateSalesTeamReport()`

**Filtering:**
- Mandatory: Company exclusion list, Sales team membership
- Optional: Revenue this FY, New clients only, Payment thresholds, Growth requirements

**Features:**
- Advanced business rules
- Growth detection logic
- Transferred account identification
- Deal-level vs company-level aggregation
- Forecasting (max 12 months)

**Output sheet:** "Sales Team Report"
**Accessed via:** Menu item 6

---

### CS Team Report.gs (14 KB)
**Purpose:** Generate CS-focused report with extended forecasting  
**Key function:** `CS_forecastRevenue()`

**Filtering:**
- Exclude: Sales Team, C-Suite
- Include: CS Team, mixed ownership

**Features:**
- Extended forecasting (no 12-month limit)
- Forecast to FY end or Report End Date
- Carry-forward logic (last payment amount)
- Separate Realized vs Forecasted columns
- Grey font for forecast months

**Output sheet:** "Summary_CS_Team"
**Accessed via:** Menu item 7

---

### Menu.gs (2.8 KB)
**Purpose:** Custom menu creation and trigger management  
**Key function:** `onOpen()` - Runs automatically on sheet open

**Menu structure:**
```
HubSpot Sync
├── 1. Sync New Deals
├── 2. Sync Missing Companies
├── 3. Sync Missing Contacts
├── 4. Update ID to Names
├── 5. Generate Summary
├── 6. Run Sales Team Report
└── 7. Run CS Team Report
```

**Auto-runs:** Every time spreadsheet opens

---

### HubSpot Company Mapper.gs (5.5 KB)
**Purpose:** TLD-based company name mapping  
**Key functions:**
- `syncAllPipelineData()` - Update all pipeline sheets
- `enrichPipelineWithNames()` - Enrich specific pipeline
- `getTLDMap()` - Get domain-to-company mappings

**Features:**
- Maps email domains to official company names
- Batch processing for performance
- Fallback to "Individual-" pattern
- Updates columns L, M, N, O

**Accessed via:** Menu item 4

---

### TLD Maping.gs (3.9 KB)
**Purpose:** Domain mapping configuration  
**Key function:** `getTLDMap()` - Returns domain-to-name mapping

**Contents:**
- Email domain mappings (e.g., '@acme.com' → 'Acme Corporation')
- Used for standardizing company names
- Extensible dictionary pattern

**Modified by:** Add new domain mappings as needed

---

### Map TLD in Summary.gs (1.9 KB)
**Purpose:** Apply TLD mapping to Summary report  
**Key function:** `applyTLDMappingToSummary()`

**Features:**
- Updates company names in Summary Report
- Uses TLD mapping dictionary
- Runs after report generation

---

## 🗂️ Repository Files

### .gitignore (4.4 KB)
**Purpose:** Specify files to exclude from version control  

**Excludes:**
- Sensitive files (tokens, credentials, secrets)
- Development files (.vscode, .idea, *.swp)
- Backup files (*.backup, *.bak)
- Temporary files (test_*.gs, temp_*.gs)
- Data exports (*.csv, *.xlsx)
- Operating system files (.DS_Store, Thumbs.db)

**Important:** Always check before committing!

---

### .git/ (Directory)
**Purpose:** Git version control metadata  
**Contents:** Git history, branches, configuration  
**Note:** Do not manually edit

---

## 📊 File Organization

### By Purpose

**Setup & Configuration:**
- README.md
- QUICKSTART.md
- Config

**Core Functionality:**
- hubspot_sync.gs
- Temp Sheet.gs

**Reporting:**
- Summary reprot.gs
- Sales Team Report.gs
- CS Team Report.gs

**Utilities:**
- Menu.gs
- HubSpot Company Mapper.gs
- TLD Maping.gs
- Map TLD in Summary.gs

**Developer Resources:**
- ARCHITECTURE.md
- CONTRIBUTING.md
- CHANGELOG.md
- LICENSE

**Project Management:**
- .gitignore
- FILE_STRUCTURE.md

---

## 📏 File Sizes Summary

| Category | Files | Total Size |
|----------|-------|------------|
| Documentation | 7 files | ~96 KB |
| Code Files | 8 files | ~92 KB |
| Config | 2 files | ~11 KB |
| **Total** | **17 files** | **~199 KB** |

---

## 🔍 Quick Reference

### Find a specific file:

**Need setup instructions?** → README.md or QUICKSTART.md  
**Need to configure settings?** → Config  
**Need to sync from HubSpot?** → hubspot_sync.gs  
**Need to understand classification?** → Temp Sheet.gs  
**Need to generate reports?** → Summary/Sales/CS Report.gs  
**Need to make changes?** → CONTRIBUTING.md  
**Need architecture details?** → ARCHITECTURE.md  
**Need version history?** → CHANGELOG.md  

---

## 🗺️ Navigation Map

```
Start Here
    ↓
README.md ──→ QUICKSTART.md (if new user)
    ↓
Config (setup token)
    ↓
Menu.gs (run sync)
    ↓
hubspot_sync.gs (imports data)
    ↓
Temp Sheet.gs (processes data)
    ↓
Summary/Sales/CS Report.gs (creates reports)
    ↓
CONTRIBUTING.md (if making changes)
```

---

## 📝 File Naming Conventions

**Documentation:** UPPERCASE.md (README.md, QUICKSTART.md)  
**Code files:** Description.gs (Menu.gs, Config)  
**Special files:** .filename (.gitignore, .git/)  

**Note:** Some files have spaces due to Google Apps Script naming:
- `CS Team Report.gs`
- `HubSpot Company Mapper.gs`
- `Map TLD in Summary.gs`
- `Sales Team Report.gs`
- `Summary reprot.gs` (typo in original, kept for compatibility)
- `TLD Maping.gs` (typo in original, kept for compatibility)
- `Temp Sheet.gs`

---

## 🚀 Getting Started Checklist

For new users:
- [ ] Read README.md overview
- [ ] Follow QUICKSTART.md steps
- [ ] Configure Config with HubSpot token
- [ ] Create required sheets
- [ ] Run first sync
- [ ] Generate reports

For developers:
- [ ] Read README.md and ARCHITECTURE.md
- [ ] Review CONTRIBUTING.md
- [ ] Understand code in hubspot_sync.gs and Temp Sheet.gs
- [ ] Set up dev environment
- [ ] Review common tasks in CONTRIBUTING.md

---

*Last Updated: February 16, 2025*
*Version: 2.0.0*
*Total Files: 17*
*Total Documentation: 96 KB*
*Total Code: 92 KB*
