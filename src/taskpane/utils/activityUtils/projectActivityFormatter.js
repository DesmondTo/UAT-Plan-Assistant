import { changeFillColor } from "../fillUtils";
import { boldFontInRange, colorFontInRange } from "../fontUtils";

/**
 * Colors the frame of the project activity.
 * @param {Excel.Range} projectActivityCell
 */
export const formatProjectActivityFrame = async (projectActivityCell) => {
  await Excel.run(async (context) => {
    await changeFillColor(projectActivityCell.getColumnsBefore(1).getRowsBelow(1), "#94C5EE");
    await changeFillColor(projectActivityCell.getEntireRow(), "#94C5EE");
    await boldFontInRange(projectActivityCell);
    await colorFontInRange(projectActivityCell, "white");
    await context.sync();
  });
};
