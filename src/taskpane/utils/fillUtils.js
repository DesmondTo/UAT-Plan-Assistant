/**
 * Changes the fill color of the specified range.
 * @param {Excel.Range} range
 * @param {string} color
 */
export const changeFillColor = async (range, fillColor) => {
  await Excel.run(async (context) => {
    range.format.fill.color = fillColor;
    await context.sync();
  });
};
