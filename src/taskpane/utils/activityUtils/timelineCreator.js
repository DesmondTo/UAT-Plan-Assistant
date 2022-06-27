import { addCalendar } from "../projectUtils/projectCalendarCreator";
import { getDayOfTheMonth, getDayDifference, getDayNumFromKickOffMonth } from "../dateUtils/dateGetter";

/**
 * Adds timeline to the selected activity in current worksheet.
 * @param {Excel.Range} activityAddress
 * @param {string} startDate
 * @param {string} endDate
 */
export const addTimeline = async (activityAddress, startDate, endDate) => {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    // Add the project calendar if it doesn't exist.
    await addCalendar(startDate, endDate);

    const date = new Date(startDate);
    const month = date.getMonth();
    const year = date.getFullYear();
    let dayNumFromFirstDayOfKickOffMonth = 0;
    await getDayNumFromKickOffMonth(year, month).then((dayNum) => {
      dayNumFromFirstDayOfKickOffMonth += dayNum;
    });

    const activityCell = currentWorksheet.getRange(activityAddress);
    activityCell.load("rowIndex");
    await context.sync();

    const startDateCell = currentWorksheet
      .getRange()
      .getRow(activityCell.rowIndex)
      .getColumn(dayNumFromFirstDayOfKickOffMonth + getDayOfTheMonth(startDate) + 3);
    startDateCell.load(["rowIndex", "columnIndex"]);
    await context.sync();

    const dayNumDiff = getDayDifference(new Date(startDate), new Date(endDate));
    const row = startDateCell.rowIndex;
    const startingCol = startDateCell.columnIndex;
    // Best practice to do split loop when loading properties.
    const cellFillArr = [];
    for (let col = 0; col <= dayNumDiff; col++) {
      const currCell = currentWorksheet.getCell(row, startingCol + col);
      currCell.load("format");
      await context.sync();
      const currCellFormat = currCell.format;
      currCellFormat.load("fill");
      await context.sync();
      cellFillArr.push(currCellFormat.fill);
    }

    cellFillArr.forEach(async (cellFill) => {
      cellFill.color = "#C7DFFA";
      await context.sync();
    });
  }).catch((error) => {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
};
