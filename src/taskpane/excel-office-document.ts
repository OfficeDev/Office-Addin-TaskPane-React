/* global Excel */

const insertText = async (text: string) => {
  // Write text to the top left cell.
  return Excel.run(async (context) => {
    const sheet: Excel.Worksheet = context.workbook.worksheets.getActiveWorksheet();
    const range: Excel.Range = sheet.getRange("A1");
    range.values = [[text]];
    return context.sync();
  });
};

export default insertText;
