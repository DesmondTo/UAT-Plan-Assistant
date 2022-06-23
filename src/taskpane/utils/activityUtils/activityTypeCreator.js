import { boldFontInRange } from "../fontUtils";

/**
 * Adds activity type to the selected cell in current worksheet.
 * @param {string} activityTypeTitle
 */
export const addActivityType = async (activityTypeTitle) => {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const currentActiveCell = context.workbook.getActiveCell();
    currentActiveCell.load(["values", "address"]);
    await context.sync();
    currentActiveCell.values = activityTypeTitle;

    await boldFontInRange(currentWorksheet.getRange(currentActiveCell.address));

    await context.sync();
  }).catch((error) => {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
};
