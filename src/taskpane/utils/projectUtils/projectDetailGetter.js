import { PROJECT_KICKOFF_DATE_CELL_INDEX } from "../../../constants/projectConstants";

import { toMonth } from "../dateUtils/dateFormatter";

/**
 * Gets the number of rows in current project worksheet.
 * @returns The number of rows in current project worksheet.
 */
export const getProjectRowCount = async () => {
  let rowCount = 0;
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const worksheetFirstColumn = currentWorksheet.getRange("A:A");
    worksheetFirstColumn.load("rowCount");
    await context.sync();

    for (let row = 0; row < worksheetFirstColumn.rowCount; row++) {
      const currCellInFirstCol = worksheetFirstColumn.getRow(row);
      currCellInFirstCol.load("format");
      await context.sync();
      const cellFormat = currCellInFirstCol.format.load("fill");
      await context.sync();
      if (cellFormat.fill.color === "#FFFFFF") {
        return;
      }
      rowCount++;
    }
  }).catch((error) => {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });

  return rowCount;
};

/**
 * Gets the date of when the project started.
 * @returns The date of when the project started with the format of YYYY-MM-DD, in string.
 */
export const getProjectKickOffDate = async () => {
  let kickOffDateStr;
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    let kickOffDate = currentWorksheet.getRange(PROJECT_KICKOFF_DATE_CELL_INDEX);
    kickOffDate.load("values");
    await context.sync();

    kickOffDate = kickOffDate.values[0][0].replace("Kick-Off Date: ", ""); // Get the date in the value array.
    const [kickOffDay, kickOffMonth, kickOffYear] = kickOffDate.split(" "); // Split the date into three entities.
    // Month need to plus 1 as toMonth return zero indexed month.
    return `${kickOffYear}-${toMonth(kickOffMonth) + 1}-${kickOffDay}`; // Month is in string, change it to number.
  }).then((dateStr) => {
    kickOffDateStr = dateStr;
  });

  return kickOffDateStr;
};
