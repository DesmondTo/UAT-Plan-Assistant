import { changeFillColor } from "../fillUtils";

/**
 * Adds activity to the selected cell in current worksheet.
 * @param {string} activityTitle
 */
export const addActivity = async (activityTitle) => {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const currentActiveCell = context.workbook.getActiveCell();
    currentActiveCell.load(["values", "address"]);
    await context.sync();
    currentActiveCell.values = activityTitle;

    const currActiveCellRange = currentWorksheet.getRange(currentActiveCell.address);
    await changeFillColor(currActiveCellRange.getColumnsBefore(1), "#94C5EE");

    await context.sync();
  }).catch((error) => {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
};
