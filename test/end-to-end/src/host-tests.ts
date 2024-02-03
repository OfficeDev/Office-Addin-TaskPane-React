import { default as insertExcelText } from "../../../src/taskpane/excel-office-document";
import { default as insertPowerPointText } from "../../../src/taskpane/powerpoint-office-document";
import { default as insertWordText } from "../../../src/taskpane/word-office-document";
import * as testHelpers from "./test-helpers";
import { sendTestResults } from "office-addin-test-helpers";

/* global document Excel Office PowerPoint Word  */

let testValues: any = [];

export const testExcelEnd2End = async (testServerPort: number): Promise<void> => {
  try {
    // Execute taskpane code
    insertExcelText("Hello Excel End2End Test");
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
  } catch {
    Promise.reject();
  }
};

export const testPowerPointEnd2End = async (testServerPort: number): Promise<void> => {
  try {
    // Execute taskpane code
    insertPowerPointText("Hello PowerPoint End2End Test");
    await testHelpers.sleep(2000);

    // Get output of executed taskpane code
    PowerPoint.run(async (context: PowerPoint.RequestContext) => {
      // get text from selected text shape
      const shapes = context.presentation.getSelectedShapes();
      const shape = shapes.getItemAt(0);
      shape.textFrame.textRange.load("text");
      await context.sync();
      const selectedText = shape.textFrame.textRange.text;

      // send test results
      testHelpers.addTestResult(testValues, "output-message", selectedText, "Hello PowerPoint End2End Test");
      await sendTestResults(testValues, testServerPort);
      testValues.pop();
      Promise.resolve();
    });
  } catch {
    Promise.reject();
  }
};

export const testWordEnd2End = async (testServerPort: number): Promise<void> => {
  try {
    // Execute taskpane code
    insertWordText("Hello Word End2End Test");
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
  } catch {
    Promise.reject();
  }
};
