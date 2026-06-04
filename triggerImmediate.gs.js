function triggerImmediate(e) {

  const sheet = e.source.getActiveSheet();

  // Only run on delivery report
  if (sheet.getName() !== "delivery report") return;

  const range = e.range;

  const startRow = range.getRow();
  const numRows = range.getNumRows();

  const col = range.getColumn();
  const numCols = range.getNumColumns();
  const lastCol = col + numCols - 1

  // N column only
  if (col > 14 || lastCol < 14) return;

  // Get pasted values
  const values = range.getValues();

  for (let i = 0; i < numRows; i++) {

    const row = startRow + i;

    // Ignore header
    if (row < 2) continue;

    const value = values[i][0];

    // Skip empty
    if (!value) continue;

    try {
      processTrackingRow(sheet, row);
    } catch (err) {
      Logger.log(err);
    }

  }

}