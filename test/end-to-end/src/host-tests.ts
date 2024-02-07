import { default as insertExcelText } from "../../../src/taskpane/excel-office-document";
import { default as insertPowerPointText } from "../../../src/taskpane/powerpoint-office-document";
import { default as insertWordText } from "../../../src/taskpane/word-office-document";
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
    // Execute taskpane code
    await testHelpers.sleep(10000);
    await insertPowerPointText("Hello PowerPoint End2End Test");
    await testHelpers.sleep(2000);

    // Get output of executed taskpane code
    await PowerPoint.run(async (context: PowerPoint.RequestContext) => {
      let selectedText = "";

      // get text from selected text shape
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items");
      await context.sync();
      if (!shapes) {
        selectedText = "No shapes object";
      } else if (shapes.items.length === 0) {
        selectedText = "No shapes selected";
      } else {
        const shape = shapes.getItemAt(0);
        const textRange = shape.textFrame.textRange.load("text");
        await context.sync();
        selectedText = textRange.text;
      }

      // send test results
      testHelpers.addTestResult(testValues, "output-message", selectedText, "Hello PowerPoint End2End Test");
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
    await testHelpers.sleep(10000);
    await insertWordText("Hello Word End2End Test");
    await testHelpers.sleep(2000);

    // Get output of executed taskpane code
    await Word.run(async (context) => {
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
  let errorMessage = "";
  if (error instanceof Error) {
    if ("message" in error) {
      errorMessage += `${error.name}: ${error.message}\n`;
    }
    if ("code" in error) {
      errorMessage += `CODE: ${error.code}\n`;
    }
    if ("stack" in error) {
      errorMessage += `STACK: ${error.stack}\n`;
    }
    if ("debugInfo" in error) {
      errorMessage += `DEBUG INFO: ${error.debugInfo}\n`;
    }
  } else {
    errorMessage = error;
  }
  return errorMessage;
};
