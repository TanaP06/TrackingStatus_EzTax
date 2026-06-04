function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚚 Tracking")
    .addItem("Update Tracking Statuses", "t2updateTrackingStatuses")
    .addToUi();
}

function t2updateTrackingStatuses() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const startRow = 2;
  const aiValues = sheet.getRange("AI:AI").getValues();

  let lastRow = 0;

  for (let i = aiValues.length - 1; i >= 0; i--) {

    if (aiValues[i][0] !== "") {

      lastRow = i + 1;
      break;

    }
  }

  // Dropdown source
  const sourceCell = sheet.getRange("AD2");

  // Read column N 14
  const values = sheet
    .getRange(startRow, 35, lastRow - startRow + 1, 1) // AI 35
    .getValues();

  values.forEach((row, index) => {
    const rowNumber = index + startRow;
    const valueInAE = sheet.getRange(rowNumber, 39).getValue(); //AD 30 AM 39
    if (valueInAE !== "") {
      return;
    }

    const oldUrl = row[0];

    // Skip empty
    if (!oldUrl) return;

    try {

      // Extract tracking_no
      const match = String(oldUrl)
        .match(/tracking_no=([^&]+)/);

      if (!match) {

        sheet.getRange(index + startRow, 36)
          .setValue("INVALID URL");

        return;
      }

      const trackingNo = match[1];

      // Create new URL directly
      const url =
        `https://aden-ali-bo.aden.asia/public/tracking?tracking_no=${trackingNo}`;

      // Fetch API
      const response = UrlFetchApp.fetch(url);

      const data = JSON.parse(response.getContentText());

      const tracking = data.data.tracking[0];
      const latest = tracking.history[0];

      const status = latest.status;
      const message = latest.message;

      let date = "";

      if (message === "Successfully delivered" && status === "DELIVERED" && latest.createdAt) {
        const raw = latest.createdAt.split("T")[0]; // handles ISO: "2025-11-21T..."
        const parsed = new Date(latest.createdAt);

        if (!isNaN(parsed)) {
          // Always format as YYYY-MM-DD
          const y = parsed.getFullYear();
          const m = String(parsed.getMonth() + 1).padStart(2, "0");
          const d = String(parsed.getDate()).padStart(2, "0");
          date = `${y}-${m}-${d}`;
        } else {
          date = raw; // fallback
        }
      }

      const aeStatus =
        message === "Successfully delivered" && status === "DELIVERED"
          ? "pending e-tax"
          : "";

      // Write AJ+2 columns together
      sheet.getRange(rowNumber, 36, 1, 3)
        .setValues([
          [status, message, date]
        ]);

      // Center align date in AL
      sheet.getRange(rowNumber, 38)
        .setHorizontalAlignment("center");

      // Copy dropdown validation from AE2
      const targetCell = sheet.getRange(`AM${rowNumber}`);

      sourceCell.copyTo(
        targetCell,
        SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
        false
      );

      // Set dropdown value
      targetCell.setValue(aeStatus);

    } catch (error) {

      const rowNumber = index + startRow;

      sheet.getRange(rowNumber, 36)
        .setValue("ERROR");

      Logger.log("Row " + rowNumber);
      Logger.log(error);

    }

  });

  SpreadsheetApp.getUi()
    .alert("Tracking statuses updated!");

}