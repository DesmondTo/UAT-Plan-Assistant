/**
 * Bolds font in the specified range.
 * @param {Excel.Range} range
 */
export const boldFontInRange = async (range) => {
  await Excel.run(async (context) => {
    range.format.font.bold = true;
    await context.sync();
  });
};

/**
 * Colors the font in the specified range.
 * @param {Excel.Range} range
 * @param {string} fontColor
 */
export const colorFontInRange = async (range, fontColor) => {
  await Excel.run(async (context) => {
    range.format.font.color = fontColor;
    await context.sync();
  });
};
