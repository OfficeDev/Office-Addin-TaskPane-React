/* global OneNote console */

const insertText = async (text: string) => {
  // Write text to the title.
  try {
    await OneNote.run(async (context) => {
      const page = context.application.getActivePage();
      page.title = text;
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
};

export default insertText;
