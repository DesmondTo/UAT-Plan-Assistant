import { addCalendar } from "../projectUtils/projectCalendarCreator";

/**
 * Adds timeline to the selected activity in current worksheet.
 * @param {string} startDate
 * @param {string} endDate
 */
export const addTimeline = async (startDate, endDate) => {
  await Excel.run(async (context) => {
    // const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const currentActiveCell = context.workbook.getActiveCell();
    currentActiveCell.load("address");
    await context.sync();
    await addCalendar(startDate, endDate);

    await context.sync();
  }).catch((error) => {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
};
