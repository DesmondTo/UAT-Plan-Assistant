import {
  PROJECT_NAME_RANGE,
  PROJECT_NAME_CELL_INDEX,
  PROJECT_HEADER_RANGE,
  PROJECT_KICKOFF_DATE_RANGE,
  PROJECT_COLUMN_HEADER_RANGE,
} from "../../../constants/projectConstants";

import {
  autoFitEntireWorksheet,
  formatProjectNameRange,
  formatProjectKickOffDateRange,
  formatProjectColumnHeaderRange,
  formatProjectHeaderRange,
} from "./projectFormatter";

import { addCalendar } from "./projectCalendarCreator";
import { toLongDate } from "../dateUtils/dateFormatter";

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
 * Initializes a project kick-off date header with proper formatting.
 * @param {Excel.Worksheet} currentWorksheet
 * @param {string} kickOffDate
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
 * Initializes a project column headers with proper formatting.
 * @param {Excel.Worksheet} currentWorksheet
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
    await addCalendar(kickOffDate, kickOffDate);

    await context.sync();
  }).catch((error) => {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
};
