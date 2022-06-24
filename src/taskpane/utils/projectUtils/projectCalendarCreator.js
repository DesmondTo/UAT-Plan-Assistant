import { autoFitRange, formatWeekdayCell } from "./projectFormatter";
import { formatMonthCalendar } from "../projectUtils/projectCalendarFormatter";
import { toShortDate, toWeekDay } from "../dateUtils/dateFormatter";
import { getDayNumFromKickOffMonth, getDateStringArrayIncreasedByMonth } from "../dateUtils/dateGetter";

/**
 * Adds the calendar of the specified date.
 * @param {string} dateStr
 */
const addProjectCalendar = async (dateStr) => {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const date = new Date(dateStr);
    const month = date.getMonth();
    const year = date.getFullYear();

    let dayNumFromFirstDayOfKickOffMonth = 0;
    await getDayNumFromKickOffMonth(year, month).then((dayNum) => {
      dayNumFromFirstDayOfKickOffMonth += dayNum;
    });

    const startingRange = currentWorksheet
      .getRange()
      .getRow(0)
      .getColumn(dayNumFromFirstDayOfKickOffMonth + 3);
    const dayNumOfTheMonth = new Date(year, month + 1, 0).getDate();
    const monthRange = startingRange.getColumnsAfter(dayNumOfTheMonth);
    monthRange.load(["values", "text"]);
    await context.sync();

    monthRange.merge();
    // Optimizes performance, if the month calendar exists, do not add.
    if (monthRange.text[0][0] === toShortDate(dateStr)) {
      await context.sync();
      return;
    }

    monthRange.values = toShortDate(dateStr);
    const monthIsEven = month % 2 === 0;
    await formatMonthCalendar(monthRange, monthIsEven);

    const firstWeekday = new Date(year, month, 1).getDay();
    await addProjectWeekdayCalendar(context, monthRange, firstWeekday);
    await addProjectDateCalendar(context, monthRange);

    await context.sync();
  }).catch((error) => {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
};

/**
 * Add the project weekday calendar for the month calendar.
 * @param {Excel.Range} monthRange The range of the month calendar.
 */
const addProjectWeekdayCalendar = async (context, monthRange, firstWeekday) => {
  await Excel.run(async () => {
    const weekdayRange = monthRange.getRowsBelow(1);
    weekdayRange.load("columnCount");
    await context.sync();

    for (let col = 0; col < weekdayRange.columnCount; col++) {
      const currColumn = weekdayRange.getColumn(col);
      currColumn.load(["format", "values"]);
      await context.sync();

      const weekdayNum = 7;
      currColumn.values = toWeekDay((firstWeekday + col) % weekdayNum);
      await formatWeekdayCell(currColumn.values, currColumn.format);
    }

    await context.sync();
  }).catch((error) => {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
};

/**
 * Add the date calendar for the month calendar.
 * @param {Excel.Range} monthRange The range of the month.
 */
const addProjectDateCalendar = async (context, monthRange) => {
  await Excel.run(async () => {
    const weekdayRange = monthRange.getRowsBelow(2).getLastRow();
    weekdayRange.load("columnCount");
    await context.sync();
    for (let col = 0; col < weekdayRange.columnCount; col++) {
      const currColumn = weekdayRange.getColumn(col);
      currColumn.load("values");
      await context.sync();

      currColumn.values = col + 1;
    }
    await autoFitRange(weekdayRange);

    await context.sync();
  });
};

/**
 * Adds the calendar from specified start date to end date.
 * @param {string} startDate
 * @param {string} endDate
 */
export const addCalendar = async (startDate, endDate) => {
  await Excel.run(async (context) => {
    // Gets an array of date string between startDate and endDate, each differs by one month.
    const dateListToAdd = getDateStringArrayIncreasedByMonth(startDate, endDate);
    dateListToAdd.forEach(async (dateStr) => {
      await addProjectCalendar(dateStr);
    });

    await context.sync();
  }).catch((error) => {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
};
