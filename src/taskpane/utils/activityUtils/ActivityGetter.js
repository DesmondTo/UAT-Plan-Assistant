/**
 * Gets all the activity for the project activity in current worksheet.
 * Each activity is a JavaScript object containing key, text and address.
 * @param projectActivityAddress The address of project activity to look through.
 * @returns An array of project activity.
 */
export const getActivityOfProjectActivity = async (projectActivityAddress) => {
  const activities = [];

  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const projectActivityRange = currentWorksheet.getRange(projectActivityAddress);

    let currCell = projectActivityRange.getRowsBelow(1);
    currCell.load(["format", "values"]);
    await context.sync();
    let currCellFormat = currCell.format;
    currCellFormat.load(["fill", "font"]);
    await context.sync();

    while (currCellFormat.fill.color === "#FFFFFF" && currCell.values[0][0] !== '') {
      currCell.load(["values", "address"]);
      await context.sync();
      if (!currCellFormat.font.bold) {
        activities.push({
          key: `${currCell.values}: ${currCell.address}`,
          text: currCell.values,
          address: currCell.address,
        });
      }
      currCell = currCell.getRowsBelow(1);
      currCell.load(["format", "values"]);
      await context.sync();

      currCellFormat = currCell.format;
      currCellFormat.load(["fill", "font"]);
      await context.sync();
    }

    await context.sync();
  }).catch((error) => {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });

  return activities;
};
