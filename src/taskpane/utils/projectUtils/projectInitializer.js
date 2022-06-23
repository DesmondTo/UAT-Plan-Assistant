import {
  PROJECT_NAME_RANGE,
  PROJECT_NAME_CELL_INDEX,
  PROJECT_HEADER_RANGE,
  PROJECT_KICKOFF_DATE_RANGE,
  PROJECT_COLUMN_HEADER_RANGE,
  PROJECT_CALENDAR_MONTH_START_INDEX_EXCLUSIVE,
  PROJECT_CALENDAR_WEEKDAY_START_INDEX_EXCLUSIVE,
  PROJECT_CALENDAR_DATE_START_INDEX_EXCLUSIVE,
} from "../../../constants/projectConstants";

import {
  autoFitEntireWorksheet,
  autoFitRange,
  formatProjectNameRange,
  formatProjectKickOffDateRange,
  formatProjectColumnHeaderRange,
  formatProjectHeaderRange,
  formatInitialMonthRange,
  formatWeekdayCell,
} from "./projectFormatter";

import { toLongDate, toShortDate, toWeekDay } from "../dateUtils/dateFormatter";

/**
 * Initializes the project month calendar from the kick-off date.
 * Only the month of the kick-off date is initialized.
 * @param {string} kickOffDate A date of when the project started.
 * @param {number} dayNums The number of days until the end of the month starting from the kick-off date.
 */
const initializeProjectMonthCalendar = async (kickOffDate, dayNums) => {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const initialMonthRange = currentWorksheet
      .getRange(PROJECT_CALENDAR_MONTH_START_INDEX_EXCLUSIVE)
      .getColumnsAfter(dayNums);

    initialMonthRange.merge();
    initialMonthRange.values = toShortDate(kickOffDate);
    await formatInitialMonthRange(currentWorksheet, initialMonthRange);

    await context.sync();
  });
};

/**
 * Initializes the project weekday calendar from the kick-off date.
 * @param {string} kickOffDate A date of when the project started.
 * @param {number} kickOffDay
 * @param {number} dayNums The number of days until the end of the month starting from the kick-off date.
 */
const initializeProjectWeekdayCalendar = async (firstWeekday, dayNums) => {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const initialWeekDayRange = currentWorksheet
      .getRange(PROJECT_CALENDAR_WEEKDAY_START_INDEX_EXCLUSIVE)
      .getColumnsAfter(dayNums);
    initialWeekDayRange.load("columnCount");
    await context.sync();

    for (let col = 0; col < initialWeekDayRange.columnCount; col++) {
      const currColumn = initialWeekDayRange.getColumn(col);
      currColumn.load(["format", "values"]);
      await context.sync();

      const weekdayNum = 7;
      currColumn.values = toWeekDay((firstWeekday + col) % weekdayNum);
      await formatWeekdayCell(currColumn.values, currColumn.format);
    }

    await context.sync();
  });
};

/**
 * Initializes the project date calendar from the kick-off date.
 * @param {string} kickOffDate A date of when the project started.
 * @param {number} kickOffDay
 * @param {number} dayNums The number of days until the end of the month starting from the kick-off date.
 */
const initializeProjectDateCalendar = async (dayNums) => {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const initialDateRange = currentWorksheet
      .getRange(PROJECT_CALENDAR_DATE_START_INDEX_EXCLUSIVE)
      .getColumnsAfter(dayNums);
    initialDateRange.load("columnCount");
    await context.sync();
    for (let col = 0; col < initialDateRange.columnCount; col++) {
      const currColumn = initialDateRange.getColumn(col);
      currColumn.load("values");
      await context.sync();

      currColumn.values = col + 1;
    }
    autoFitRange(initialDateRange);

    await context.sync();
  });
};

const initializeProjectCalendar = async (kickOffDate) => {
  await Excel.run(async (context) => {
    const kickOffDateObj = new Date(kickOffDate);
    const kickOffYear = kickOffDateObj.getFullYear();
    const kickOffMonth = kickOffDateObj.getMonth() + 1;
    const firstWeekday = new Date(kickOffYear, kickOffMonth, 1).getDay();
    const dayNums = new Date(kickOffYear, kickOffMonth, 0).getDate();

    await initializeProjectMonthCalendar(kickOffDate, dayNums);
    await initializeProjectWeekdayCalendar(firstWeekday, dayNums);
    await initializeProjectDateCalendar(dayNums);

    await context.sync();
  }).catch((error) => {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
};

/**
 * Initializes a project name header with proper formatting.
 * @param {Excel.Worksheet} currentWorksheet
 * @param {string} projectName
 */
const initializeProjectNameHeader = async (currentWorksheet, projectName) => {
  await Excel.run(async (context) => {
    const projectNameRange = currentWorksheet.getRange(PROJECT_NAME_RANGE);
    await formatProjectNameRange(projectNameRange);
    currentWorksheet.getRange(PROJECT_NAME_CELL_INDEX).values = `Project: ${projectName}`;

    await context.sync();
  });
};

/**
 * Initializes a project name header with proper formatting.
 * @param {Excel.Worksheet} currentWorksheet
 * @param {string} projectName
 */
const initializeProjectKickOffDateHeader = async (currentWorksheet, kickOffDate) => {
  await Excel.run(async (context) => {
    const projectKickOffDateRange = currentWorksheet.getRange(PROJECT_KICKOFF_DATE_RANGE);
    await formatProjectKickOffDateRange(projectKickOffDateRange);
    projectKickOffDateRange.values = [[`Kick-Off Date: ${toLongDate(kickOffDate)}`], [""]];

    await context.sync();
  });
};

/**
 * Initializes a project name header with proper formatting.
 * @param {Excel.Worksheet} currentWorksheet
 * @param {string} projectName
 */
const initializeProjectColumnHeaders = async (currentWorksheet) => {
  await Excel.run(async (context) => {
    const projectColumnHeaderRange = currentWorksheet.getRange(PROJECT_COLUMN_HEADER_RANGE);
    projectColumnHeaderRange.values = [
      ["Status", "Action Party"],
      ["", ""],
    ];

    await formatProjectColumnHeaderRange(projectColumnHeaderRange);

    await context.sync();
  });
};

/**
 * Initializes a project plan template.
 * @param {string} projectName
 * @param {string} kickOffDate A date of when the project started.
 */
export const initializeProject = async (projectName, kickOffDate) => {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.getRange().clear(); /* Clear all the values and formats in the currect worksheet */

    await initializeProjectNameHeader(currentWorksheet, projectName);
    await initializeProjectKickOffDateHeader(currentWorksheet, kickOffDate);
    await initializeProjectColumnHeaders(currentWorksheet);
    await autoFitEntireWorksheet(currentWorksheet);

    currentWorksheet.freezePanes.freezeAt(PROJECT_HEADER_RANGE);
    currentWorksheet.getCell(0, 0).format.columnWidth = 20; /* Makes the first column narrower */
    const projectHeaderRange = currentWorksheet.getRange(PROJECT_HEADER_RANGE);
    projectHeaderRange.load("columnCount");
    await context.sync();

    await formatProjectHeaderRange(context, projectHeaderRange);
    await initializeProjectCalendar(kickOffDate);
    await context.sync();
  }).catch((error) => {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
};
