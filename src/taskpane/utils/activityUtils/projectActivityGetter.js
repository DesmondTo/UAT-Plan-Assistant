import { getProjectRowCount } from "../projectUtils/projectDetailGetter";

/**
 * Gets all the project activity in current worksheet.
 * Each project activity is a JavaScript object containing title and address.
 * @returns An array of project activity.
 */
export const getAllProjectActivity = async () => {
  const allProjectActivity = [];

  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const projectActivityColumns = currentWorksheet.getRange("B:B");
    const projectRowCount = await getProjectRowCount();

    for (let row = 0; row < projectRowCount; row++) {
      const currRow = projectActivityColumns.getRow(row);
      currRow.load("values");
      await context.sync();

      const currRowValue = currRow.values[0][0]; // To get the first cell in the range.
      if (currRowValue.startsWith("Project Activity:")) {
        const shortenedText = currRowValue.replace("Project Activity:", "");
        currRow.load(["rowIndex", "columnIndex"]);
        await context.sync();
        allProjectActivity.push({
          key: `${shortenedText}: (${currRow.rowIndex}, ${currRow.columnIndex})`,
          text: shortenedText,
          rowIndex: currRow.rowIndex,
          columnIndex: currRow.columnIndex,
        });
      }
    }

    await context.sync();
  }).catch((error) => {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });

  return allProjectActivity;
};
