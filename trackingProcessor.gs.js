function processTrackingRow(sheet, rowNumber) {

  const sourceCell = sheet.getRange("AD2");

  // READ ONCE (A:AD) 
  const rowValues = sheet.getRange(rowNumber, 1, 1, 30).getValues()[0];

  const oldUrl = String(rowValues[13] || "").trim(); // N (14)
  const valueInAD = rowValues[29]; // AD (30)

  // Skip if AD already has value
  if (valueInAD !== "") return;

  // 1. Skip if URL empty
  if (!oldUrl || oldUrl === "") return;

  let statusResult = "";
  let dateResult = "";
  let messageResult = "";
  let adValue = "";

  function writeBack() {
    sheet.getRange(rowNumber, 27, 1, 2).setValues([
      [statusResult, dateResult]
    ]);

    sheet.getRange(rowNumber, 28)
      .setHorizontalAlignment("center");
  }

  try {

    const match = oldUrl.match(/tracking_no=([^&]+)/);

    // 2. If the URL doesn't contain 'tracking_no=', mark is as INVALID URL
    if (!match) {
      statusResult = "INVALID URL";
      messageResult = "";
      writeBack();
      return;
    }

    const trackingNo = match[1];

    const url =
      `https://aden-ali-bo.aden.asia/public/tracking?tracking_no=${trackingNo}`;

    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());

    if (!data?.data?.tracking?.length) {
      statusResult = "No tracking data found - check URL or tracking number";
      messageResult = "";
      writeBack();
      return;
    }

    const tracking = data.data.tracking[0];
    const latest = tracking?.history?.[0];

    if (!latest) {
      statusResult = "No tracking history found";
      messageResult = "";
      writeBack();
      return;
    }

    const status = latest.status || "N/A";
    const message = latest.message || "N/A";

    statusResult = status;
    messageResult = message;

    if (message === "Successfully delivered" && status === "DELIVERED" && latest.createdAt) {

      const parsed = new Date(latest.createdAt);

      if (!isNaN(parsed)) {
        const y = parsed.getFullYear();
        const m = String(parsed.getMonth() + 1).padStart(2, "0");
        const d = String(parsed.getDate()).padStart(2, "0");
        dateResult = `${y}-${m}-${d}`;
      } else {
        // fallback: raw ISO date part
        dateResult = String(latest.createdAt).split("T")[0];
      }
    }

    adValue =
      (message === "Successfully delivered" && status === "DELIVERED")
        ? "pending e-tax"
        : "";

    writeBack();

    // AI COLUMN (MESSAGE) 
    sheet.getRange(rowNumber, 35).setValue(messageResult); // AI

    // AD dropdown + value 
    const targetCell = sheet.getRange(`AD${rowNumber}`);

    sourceCell.copyTo(
      targetCell,
      SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
      false
    );

    targetCell.setValue(adValue);

  } catch (error) {

    // 3. If the fetch completely fails because the URL structure breaks the HTTP request
    statusResult = "Tracking request failed";
    messageResult = "";

    writeBack();

    sheet.getRange(rowNumber, 35).setValue("Tracking request failed");

    Logger.log("Row " + rowNumber);
    Logger.log(error);
  }
}