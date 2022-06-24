import {
  PROJECT_NAME_RANGE_FILL_COLOR,
  PROJECT_NAME_RANGE_FONT_COLOR,
  PROJECT_COLUMN_HEADER_RANGE_FILL_COLOR,
  PROJECT_COLUMN_HEADER_RANGE_FONT_COLOR,
} from "../../../constants/projectConstants";

/**
 * Auto fits entire worksheet.
 * @param {Excel.Worksheet} currentWorksheet
 */
export const autoFitEntireWorksheet = async (currentWorksheet) => {
  await Excel.run(async (context) => {
    currentWorksheet.getRange().format.autofitColumns();
    currentWorksheet.getRange().format.autofitRows();

    await context.sync();
  });
};

/**
 * Auto fits specified range.
 * @param {Excel.Range} range
 */
export const autoFitRange = async (range) => {
  await Excel.run(async (context) => {
    range.format.autofitColumns();
    range.format.autofitRows();

    await context.sync();
  });
};

/**
 * Formats the fill and font in the project name range.
 * @param {Excel.Range} projectNameRange
 */
export const formatProjectNameRange = async (projectNameRange) => {
  await Excel.run(async (context) => {
    projectNameRange.format.fill.color = PROJECT_NAME_RANGE_FILL_COLOR;
    projectNameRange.format.font.bold = true;
    projectNameRange.format.font.color = PROJECT_NAME_RANGE_FONT_COLOR;
    projectNameRange.format.horizontalAlignment = "Center";

    await context.sync();
  });
};

/**
 * Formats the fill and font in the project kick-off date header.
 * @param {Excel.Range} projectNameRange
 */
export const formatProjectKickOffDateRange = async (projectKickOffDateRange) => {
  await Excel.run(async (context) => {
    projectKickOffDateRange.format.fill.color = "white";
    projectKickOffDateRange.format.font.color = "black";

    await context.sync();
  });
};

/**
 * Formats the fill and font in the project column header range.
 * @param {Excel.Range} projectColumnHeaderRange
 */
export const formatProjectColumnHeaderRange = async (projectColumnHeaderRange) => {
  await Excel.run(async (context) => {
    projectColumnHeaderRange.format.fill.color = PROJECT_COLUMN_HEADER_RANGE_FILL_COLOR;
    projectColumnHeaderRange.format.font.bold = true;
    projectColumnHeaderRange.format.font.color = PROJECT_COLUMN_HEADER_RANGE_FONT_COLOR;
    projectColumnHeaderRange.format.horizontalAlignment = "Center";

    await context.sync();
  });
};

/**
 * Formats the fill and font in the project header range.
 * @param {Excel.Range} projectHeaderRange
 */
export const formatProjectHeaderRange = async (context, projectHeaderRange) => {
  await Excel.run(async () => {
    let headerColumnsFormat = [];
    for (let col = 0; col < projectHeaderRange.columnCount; col++) {
      var currColumnFormat = projectHeaderRange.getColumn(col).format;
      currColumnFormat.load("columnWidth");
      headerColumnsFormat.push(currColumnFormat);
    }

    await context.sync();

    for (let i = 0; i < headerColumnsFormat.length; i++) {
      headerColumnsFormat[i].columnWidth = headerColumnsFormat[i].columnWidth * 1.5;
    }

    await context.sync();
  }).catch((error) => {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
};
