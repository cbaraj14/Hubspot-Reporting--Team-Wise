/**
 * ============================================================================
 * CUSTOM MENU & TRIGGERS
 * ============================================================================
 * 
 * This module creates the custom "HubSpot Sync" menu that appears in the
 * Google Sheets toolbar when the spreadsheet is opened.
 * 
 * MENU STRUCTURE:
 * 
 * HubSpot Sync
 * ├── 1. Sync New Deals              → syncHubSpotDeals()
 * ├── ─────────────────
 * ├── 2. Sync Missing Companies      → syncMissingCompanies()
 * ├── 3. Sync Missing Contacts       → syncMissingContacts()
 * ├── ─────────────────
 * ├── 4. Update ID to Names          → syncAllPipelineData()
 * ├── ─────────────────
 * ├── 5. Generate Summary            → generateSummaryReport()
 * ├── 6. Run Sales Team Report       → salesGenerateSalesTeamReport()
 * └── 7. Run CS Team Report          → generateCSReport()
 * 
 * RECOMMENDED WORKFLOW:
 * 
 * First-Time Setup:
 *   1. Sync New Deals (imports all deals from HubSpot)
 *   2. Sync Missing Companies (enriches with company associations)
 *   3. Sync Missing Contacts (enriches with contact associations)
 *   4. Update ID to Names (converts IDs to readable names)
 *   5. Generate Summary (creates comprehensive report)
 * 
 * Daily/Weekly Updates:
 *   1. Sync New Deals (gets only new/updated deals)
 *   4. Update ID to Names (refreshes names)
 *   5-7. Generate desired reports
 * 
 * AUTOMATIC TRIGGERS:
 * - onOpen() runs automatically when spreadsheet is opened
 * - Creates the custom menu
 * - No authentication required after first authorization
 * 
 * ERROR HANDLING:
 * - Menu creation failures are logged but don't block opening
 * - Individual function errors are shown in UI alerts
 * 
 * See README.md for detailed usage instructions.
 * ============================================================================
 */

function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('HubSpot Sync')
      .addItem('1. Sync New Deals', 'syncHubSpotDeals')
      .addSeparator()
      .addItem('2. Sync Missing Companies', 'syncMissingCompanies')
      .addItem('3. Sync Missing Contacts', 'syncMissingContacts')
      .addSeparator()
      .addItem('4. Update ID to Names', 'syncAllPipelineData')
      .addSeparator()
      .addItem('5. Generate Summary', 'generateSummaryReport')
      .addItem('6. Run Sales Team Report', 'salesGenerateSalesTeamReport')
      .addItem('7. Run CS Team Report', 'CS_forecastRevenue')
      .addSeparator()
      //.addItem('3. Sync Missing Contacts', 'syncMissingContacts')
      .addToUi();
  } catch (e) {
    console.log('Menu creation skipped: ' + e.message);
  }
}
