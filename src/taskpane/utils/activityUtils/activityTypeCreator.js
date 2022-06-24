import { boldFontInRange, colorFontInRange } from "../fontUtils";
import { changeFillColor } from "../fillUtils";

/**
 * Adds activity type to the selected cell in current worksheet.
 * @param {string} activityTypeTitle
 */
export const addActivityType = async (activityTypeTitle, projectActivityAddress) => {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const cellBelowProjectActivityHeader = currentWorksheet.getRange(projectActivityAddress).getRowsBelow(1);
    cellBelowProjectActivityHeader.load(["values", "address"]);
    await context.sync();

    let newRow = cellBelowProjectActivityHeader.getEntireRow();
    if (cellBelowProjectActivityHeader.values[0][0] !== "") {
      // Clears the fill brought down from the project activity header.
      newRow = newRow.insert(Excel.InsertShiftDirection.down);
      newRow.format.fill.clear();
      const firstCol = newRow.getColumn(0);
      await changeFillColor(firstCol, "#94C5EE");
    }

    // Did not reuse {cellBelowProjectActivityHeader} as its property affected after insert.
    const projectActivityCell = currentWorksheet.getRange(projectActivityAddress).getRowsBelow(1);
    projectActivityCell.values = activityTypeTitle;
    await colorFontInRange(projectActivityCell, "black");
    await boldFontInRange(projectActivityCell);
    await changeFillColor(projectActivityCell.getColumnsBefore(1), "#94C5EE");

    await context.sync();
  }).catch((error) => {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
};
