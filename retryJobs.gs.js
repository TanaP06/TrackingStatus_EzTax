function retryPendingTracking() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("delivery report");

  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Batch read columns:
  // N = 14 (tracking URL)
  // AD = 30 (status)
  const nValues = sheet.getRange(2, 14, lastRow - 1, 1).getValues();
  const adValues = sheet.getRange(2, 30, lastRow - 1, 1).getValues();

  for (let i = 0; i < lastRow - 1; i++) {

    const n = nValues[i][0];
    const ad = String(adValues[i][0]).trim();

    // Only retry rows with URL AND empty AD
    if (n && ad === "") {
      processTrackingRow(sheet, i + 2);
    }
  }
}