/* global Word */

const insertText = async (text: string) => {
  // Write text to the document.
  return Word.run(async (context) => {
    let body = context.document.body;
    body.insertParagraph(text, Word.InsertLocation.end);
    return context.sync();
  });
};

export default insertText;
