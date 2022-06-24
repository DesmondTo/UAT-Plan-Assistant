import {
  PROJECT_CALENDAR_CELL_COLUMN_WIDTH,
  PROJECT_CALENDAR_MONTH_CALENDAR_FILL_COLOR_DARK,
  PROJECT_CALENDAR_MONTH_CALENDAR_FILL_COLOR_LIGHT,
  PROJECT_SATURDAY_CALENDAR_CELL_FILL_COLOR,
  PROJECT_SUNDAY_CALENDAR_CELL_FILL_COLOR,
  PROJECT_SHORT_DATE_FORMAT,
} from "../../../constants/projectConstants";

/**
 * Formats the month calendar of the project.
 * @param {Excel.Worksheet} currentWorksheet
 * @param {Excel.Range} monthRange
 */
export const formatMonthCalendar = async (monthRange, monthIsEven) => {
  await Excel.run(async (context) => {
    monthRange.load(["numberFormat", "format"]);
    await context.sync();
    const monthRangeFormat = monthRange.format;

    monthRangeFormat.fill.color = monthIsEven
      ? PROJECT_CALENDAR_MONTH_CALENDAR_FILL_COLOR_DARK
      : PROJECT_CALENDAR_MONTH_CALENDAR_FILL_COLOR_LIGHT;
    monthRange.numberFormat = PROJECT_SHORT_DATE_FORMAT;
    monthRangeFormat.font.bold = true;
    monthRangeFormat.horizontalAlignment = "Center";
    await context.sync();
  }).catch((error) => {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
};

/**
 * Formats the weekday calendar cell of the project.
 * @param {Excel.Range.values} currColumnValue
 * @param {Excel.RangeFormat} currColumnFormat
 */
export const formatWeekdayCell = async (currColumnValue, currColumnFormat) => {
  await Excel.run(async (context) => {
    currColumnFormat.horizontalAlignment = "Center";

    if (currColumnValue == "Sa") {
      currColumnFormat.fill.color = PROJECT_SATURDAY_CALENDAR_CELL_FILL_COLOR;
    }
    if (currColumnValue == "Su") {
      currColumnFormat.fill.color = PROJECT_SUNDAY_CALENDAR_CELL_FILL_COLOR;
    }

    await context.sync();
  });
};

/**
 * Formats the date cell of the month calendar of the project.
 * @param {Excel.RangeFormat} currColumnFormat
 */
export const formatDateCell = async (currColumnFormat) => {
  await Excel.run(async (context) => {
    currColumnFormat.horizontalAlignment = "Center";
    currColumnFormat.columnWidth = PROJECT_CALENDAR_CELL_COLUMN_WIDTH;
    await context.sync();
  });
};
