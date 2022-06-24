import {
  PROJECT_CALENDAR_MONTH_CALENDAR_FILL_COLOR_DARK,
  PROJECT_CALENDAR_MONTH_CALENDAR_FILL_COLOR_LIGHT,
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
