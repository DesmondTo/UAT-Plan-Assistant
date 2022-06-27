/**
 * Chooses to bold or unbold font in the specified range.
 * @param {Excel.Range} range
 * @param {boolean} isBold
 */
export const boldFontInRange = async (range, isBold=true) => {
  await Excel.run(async (context) => {
    range.format.font.bold = isBold;
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
