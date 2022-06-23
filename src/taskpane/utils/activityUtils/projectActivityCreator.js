import { formatProjectActivityFrame } from "./projectActivityFormatter";

/**
 * Add a project activity with proper styling.
 * @param {string} projectActivityTitle
 */
export const addProjectActivity = async (projectActivityTitle) => {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    // Find the first blank cell in column A
    const entireWorksheetRange = currentWorksheet.getRange();
    const worksheetFirstColumn = entireWorksheetRange.getColumn(0);
    worksheetFirstColumn.load("rowCount");
    await context.sync();

    for (let row = 0; row < worksheetFirstColumn.rowCount; row++) {
      const currCellInFirstCol = worksheetFirstColumn.getRow(row);
      currCellInFirstCol.load("format");
      await context.sync();
      const cellFormat = currCellInFirstCol.format.load("fill");
      await context.sync();
      if (cellFormat.fill.color === "#FFFFFF") {
        const projectActivityCell = currCellInFirstCol.getColumnsAfter(1);
        projectActivityCell.values = `Project Activity: ${projectActivityTitle}`;
        await formatProjectActivityFrame(projectActivityCell);
        await context.sync();
        return;
      }
    }

    await context.sync();
  }).catch((error) => {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
};
