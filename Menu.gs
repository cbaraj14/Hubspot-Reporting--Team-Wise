// ============================================================================
// TRIGGER & MENU
// ============================================================================

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
      .addItem('7. Run CS Team Report', 'generateCSReport')
      .addSeparator()
      //.addItem('3. Sync Missing Contacts', 'syncMissingContacts')
      .addToUi();
  } catch (e) {
    console.log('Menu creation skipped: ' + e.message);
  }
}
