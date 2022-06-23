/**
 * Gets the number of rows in current project worksheet.
 * @returns The number of rows in current project worksheet.
 */
export const getProjectRowCount = async () => {
  let rowCount = 0;
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const worksheetFirstColumn = currentWorksheet.getRange("A:A");
    worksheetFirstColumn.load("rowCount");
    await context.sync();

    for (let row = 0; row < worksheetFirstColumn.rowCount; row++) {
      const currCellInFirstCol = worksheetFirstColumn.getRow(row);
      currCellInFirstCol.load("format");
      await context.sync();
      const cellFormat = currCellInFirstCol.format.load("fill");
      await context.sync();
      if (cellFormat.fill.color === "#FFFFFF") {
        return;
      }
      rowCount++;
    }
  }).catch((error) => {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });

  return rowCount;
};
