/* global OneNote */

const insertText = async (text: string) => {
  // Write text to the title.
  return OneNote.run(async (context) => {
    const page = context.application.getActivePage();
    page.title = text;
    return context.sync();
  });
};

export default insertText;
