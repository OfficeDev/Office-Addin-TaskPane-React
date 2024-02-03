/* global Word */

const insertText = async (text: string) => {
  // Write text to the document.
  await Word.run(async (context) => {
    let body = context.document.body;
    body.insertParagraph(text, Word.InsertLocation.end);
    await context.sync();
  });
  Promise.resolve();
};

export default insertText;
