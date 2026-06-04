function triggerImmediate(e) {

  const sheet = e.source.getActiveSheet();

  if (sheet.getName() !== "delivery report") return;

  const range = e.range;

  const startRow = range.getRow();
  const numRows = range.getNumRows();

  const startCol = range.getColumn();
  const endCol = startCol + range.getNumColumns() - 1;

  // Check if column N (14) is within the pasted range
  if (startCol > 14 || endCol < 14) return;

  // Get only column N values
  const colNIndex = 14 - startCol; // relative index within the range

  const values = range.getValues();

  for (let i = 0; i < numRows; i++) {

    const row = startRow + i;

    if (row < 2) continue;

    const value = values[i][colNIndex];

    if (!value) continue;

    try {
      processTrackingRow(sheet, row);
    } catch (err) {
      Logger.log(err);
    }

  }

}