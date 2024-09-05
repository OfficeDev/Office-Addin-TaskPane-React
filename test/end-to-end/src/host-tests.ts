import { insertText as insertExcelText } from "../../../src/taskpane/excel";
import { insertText as insertPowerPointText } from "../../../src/taskpane/powerpoint";
import { insertText as insertWordText } from "../../../src/taskpane/word";
import * as testHelpers from "./test-helpers";
import { sendTestResults } from "office-addin-test-helpers";

/* global Excel PowerPoint Word  */

let testValues: any = [];

export const testExcelEnd2End = async (testServerPort: number): Promise<void> => {
  try {
    // Execute taskpane code
    await insertExcelText("Hello Excel End2End Test");
    await testHelpers.sleep(2000);

    // Get output of executed taskpane code
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");
      await context.sync();
      await testHelpers.sleep(2000);

      // send test results
      testHelpers.addTestResult(testValues, "output-message", range.values[0][0], "Hello Excel End2End Test");
      await sendTestResults(testValues, testServerPort);
      testValues.pop();
      await testHelpers.closeWorkbook();
      Promise.resolve();
    });
  } catch (error) {
    testHelpers.addTestResult(testValues, "output-message", getErrorMessage(error), "");
    await sendTestResults(testValues, testServerPort);
    testValues.pop();
    Promise.reject();
  }
};

export const testPowerPointEnd2End = async (testServerPort: number): Promise<void> => {
  try {
    const textToInsert = "Hello PowerPoint End2End Test";

    // Execute taskpane code
    await insertPowerPointText(textToInsert);
    await testHelpers.sleep(2000);

    // Get output of executed taskpane code
    PowerPoint.run(async (context: PowerPoint.RequestContext) => {
      // get text from inserted text shape
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      // eslint-disable-next-line office-addins/load-object-before-read, office-addins/call-sync-before-read
      const shapes = slide.shapes;
      slide.shapes.load(["textFrame/textRange/text"]);
      await context.sync();
      const shape = shapes.items[shapes.items.length - 1];
      const text = shape.textFrame.textRange.text;

      // send test results
      testHelpers.addTestResult(testValues, "output-message", text, textToInsert);
      await sendTestResults(testValues, testServerPort);
      testValues.pop();
      Promise.resolve();
    });
  } catch (error) {
    testHelpers.addTestResult(testValues, "output-message", getErrorMessage(error), "");
    await sendTestResults(testValues, testServerPort);
    testValues.pop();
    Promise.reject();
  }
};

export const testWordEnd2End = async (testServerPort: number): Promise<void> => {
  try {
    // Execute taskpane code
    await insertWordText("Hello Word End2End Test");
    await testHelpers.sleep(2000);

    // Get output of executed taskpane code
    Word.run(async (context) => {
      var firstParagraph = context.document.body.paragraphs.getFirst();
      firstParagraph.load("text");
      await context.sync();
      await testHelpers.sleep(2000);

      // send test results
      testHelpers.addTestResult(testValues, "output-message", firstParagraph.text, "Hello Word End2End Test");
      await sendTestResults(testValues, testServerPort);
      testValues.pop();
      Promise.resolve();
    });
  } catch (error) {
    testHelpers.addTestResult(testValues, "output-message", getErrorMessage(error), "");
    await sendTestResults(testValues, testServerPort);
    testValues.pop();
    Promise.reject();
  }
};

const getErrorMessage = (error: any): string => {
  if (error instanceof Error) {
    if ("stack" in error) {
      return error.stack;
    } else {
      return `${error.name}: ${error.message}`;
    }
  } else {
    return error;
  }
};
