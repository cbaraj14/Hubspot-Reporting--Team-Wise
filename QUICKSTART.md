# Quick Start Guide

Get up and running with the HubSpot Deals to Google Sheets Importer in 15 minutes.

---

## 📋 Prerequisites Checklist

Before you begin, make sure you have:

- [ ] Google Sheets access
- [ ] HubSpot CRM account with admin access
- [ ] Authorization to create HubSpot Private Apps
- [ ] 15 minutes of uninterrupted time

---

## 🚀 5-Step Setup

### Step 1: Get Your HubSpot API Token (5 min)

1. **Go to HubSpot Settings**
   - Click the ⚙️ icon in your HubSpot account
   - Navigate to **Integrations** → **Private Apps**

2. **Create a New Private App**
   - Click "Create a private app"
   - Name: `Google Sheets Integration`
   - Description: `Read-only access for deal data export`

3. **Configure Scopes**
   - Under "Scopes", check these boxes:
     - ✅ `crm.objects.deals.read`
     - ✅ `crm.objects.companies.read`
     - ✅ `crm.objects.contacts.read`

4. **Generate Token**
   - Click "Create app"
   - Copy the token (starts with `pat-na1-...`)
   - ⚠️ **Save it securely** - you won't see it again!

### Step 2: Set Up Your Google Sheet (3 min)

1. **Create Required Sheets**
   
   Create these sheets in your Google Sheets document (exact names):
   
   ```
   Config_sheet
   Sales_team_Members
   CS_team_Members
   Company_Stop_list
   ```

2. **Configure Config_sheet**
   
   Add these values:
   
   | Cell | Label (Column D) | Value (Column E) |
   |------|------------------|------------------|
   | E6 | Report Date | `6/30/2025` (or today's date) |
   | E8 | Report Start Date | `1/1/2024` |
   | E9 | Report End Date | `6/30/2026` |
   | E12 | Sales: Revenue This FY | `FALSE` |
   | E13 | Sales: Aggregation Level | `Company Level` |
   | E14 | Sales: Client Age Filter | `All` |
   | E15 | Sales: Payment Threshold | `5` |
   | E16 | Sales: Growth Check | `FALSE` |
   | E17 | Sales: Transferred Window | `12` |

3. **Set Up Team Sheets**
   
   **Sales_team_Members** sheet:
   ```
   Row 1: Owner ID  (header)
   Row 2+: Paste HubSpot Owner IDs (one per row)
   ```
   
   **CS_team_Members** sheet:
   ```
   Row 1: Owner ID  (header)
   Row 2+: Paste HubSpot Owner IDs (one per row)
   ```
   
   **Company_Stop_list** sheet:
   ```
   Row 1: Company Name  (header)
   Row 2+: Companies to exclude from Sales report (optional, can be empty)
   ```

### Step 3: Install the Scripts (4 min)

1. **Open Apps Script Editor**
   - In your Google Sheet, click **Extensions** → **Apps Script**

2. **Create Script Files**
   
   Create these files (click ➕ next to "Files"):
   - `Config.gs`
   - `hubspot_sync.gs`
   - `Temp Sheet.gs`
   - `Summary reprot.gs`
   - `Sales Team Report.gs`
   - `CS Team Report.gs`
   - `Menu.gs`

3. **Copy Code**
   - Copy the contents from the provided `.gs` files in the repository
   - Paste into corresponding files in Apps Script editor

4. **Add Your HubSpot Token**
   - Open `Config.gs`
   - Find line: `HUBSPOT_TOKEN: 'your-hubspot-token-here',`
   - Replace `'your-hubspot-token-here'` with your actual token
   - Example: `HUBSPOT_TOKEN: 'pat-na1-xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx',`

5. **Save Everything**
   - Click 💾 **Save project**
   - Give it a name: `HubSpot Integration`

### Step 4: First Sync (2 min)

1. **Reload Your Google Sheet**
   - Close the Apps Script tab
   - Refresh your Google Sheet
   - You should see a new menu: **HubSpot Sync**

2. **Authorize the Script**
   - Click **HubSpot Sync** → **1. Sync New Deals**
   - Click **Continue** when prompted
   - Select your Google account
   - Click **Advanced** → **Go to HubSpot Integration (unsafe)**
   - Click **Allow**

3. **Run First Import**
   - Click **HubSpot Sync** → **1. Sync New Deals**
   - Wait for completion (may take 1-5 minutes depending on deal count)
   - You'll see a success message

### Step 5: Generate Reports (1 min)

Run these menu items in order:

1. **HubSpot Sync** → **2. Sync Missing Companies**
2. **HubSpot Sync** → **3. Sync Missing Contacts**
3. **HubSpot Sync** → **4. Update ID to Names**
4. **HubSpot Sync** → **5. Generate Summary**

You should now see a new sheet: **Summary Report** with your deal data!

---

## ✅ Verification

Check that you have these sheets:

- [x] **Sales Pipeline** (raw HubSpot data)
- [x] **CS Pipeline** (raw HubSpot data)
- [x] **Payment Pipeline** (raw HubSpot data)
- [x] **TEMP_DATA** (hidden - pre-computed cache)
- [x] **Summary Report** (your first report! 🎉)

---

## 🎯 What's Next?

### Daily/Weekly Updates

1. **HubSpot Sync** → **1. Sync New Deals**
2. **HubSpot Sync** → **4. Update ID to Names**
3. **HubSpot Sync** → **5. Generate Summary**

### Generate Additional Reports

**Sales Team Report:**
```
HubSpot Sync → 6. Run Sales Team Report
```

**CS Team Report:**
```
HubSpot Sync → 7. Run CS Team Report
```

### Customize Filters

Edit values in **Config_sheet** cells E12-E17 to adjust Sales report filters.

---

## 🔧 Troubleshooting

### "Sheet not found" Error
**Solution:** Create the missing sheet with the exact name shown in the error.

### "Invalid credentials" Error
**Solution:** 
1. Verify your HubSpot token is correct in `Config.gs`
2. Check token hasn't expired (HubSpot tokens can be deactivated)
3. Confirm scopes are correct (deals, companies, contacts - all read)

### No Data Imported
**Solution:**
1. Check your HubSpot deals are in "Closed Won" stages
2. Verify stage IDs in `Config.gs` match your HubSpot stages
3. Look at Apps Script logs: **Extensions** → **Apps Script** → **Executions**

### Import is Very Slow
**Solution:**
1. First import is always slower (imports all historical deals)
2. Subsequent syncs only import new/modified deals
3. Consider increasing `BATCH_SIZE` in `Config.gs` from 2500 to 5000

### Missing Company Names
**Solution:** Run **HubSpot Sync** → **2. Sync Missing Companies**

---

## 📞 Getting Help

1. **Check the logs:**
   - Extensions → Apps Script → Executions
   - Look for error messages

2. **Read the full docs:**
   - See `README.md` for comprehensive documentation

3. **Common issues:**
   - Rate limiting: Increase `COOLDOWN_MS` in Config.gs
   - Timeouts: Reduce `BATCH_SIZE` in Config.gs
   - Missing data: Re-run sync functions in order (1→2→3→4)

---

## 🎓 Understanding the System

### Data Flow
```
HubSpot CRM
    ↓
[1. Sync New Deals]
    ↓
Pipeline Sheets (Sales/CS/Payment)
    ↓
[2. Sync Missing Companies]
[3. Sync Missing Contacts]
    ↓
[4. Update ID to Names]
    ↓
TEMP_DATA (hidden cache)
    ↓
[5-7. Generate Reports]
    ↓
Summary / Sales Team / CS Team Reports
```

### Key Concepts

**Pipeline Sheets:**
- Raw data from HubSpot (15 columns)
- Three sheets: Sales Pipeline, CS Pipeline, Payment Pipeline
- Payment Pipeline = source of truth for revenue

**TEMP_DATA Sheet:**
- Hidden sheet with 25+ columns
- Pre-computed classifications (Revenue Type, Client Age, etc.)
- Makes reports run 10-20x faster

**Reports:**
- Summary: All companies, no filtering
- Sales Team: Filtered for sales metrics
- CS Team: Filtered for CS metrics + forecasting

---

## 🎉 Success!

You now have a fully automated HubSpot-to-Google Sheets integration!

**Next Steps:**
- Explore the full README.md
- Customize report filters
- Set up automated triggers (optional)
- Share reports with your team

**Pro Tips:**
- Run syncs daily/weekly to keep data fresh
- Use filters in reports instead of re-syncing
- Archive old data if exceeding 10,000 rows
- Monitor HubSpot API usage in HubSpot dashboard

---

*Last Updated: February 16, 2025*
