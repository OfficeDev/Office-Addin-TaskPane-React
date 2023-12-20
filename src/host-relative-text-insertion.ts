// import { default as insertExcelText } from "./taskpane/excel-office-document";
// import { default as insertPowerPointText } from "./taskpane/powerpoint-office-document";
// import { default as insertWordText } from "./taskpane/word-office-document";
// import { default as insertOutlookText } from "./taskpane/outlook-office-document";
import { insertWordText, insertExcelText, insertPowerPointText, insertOutlookText } from "./taskpane/office-document";

/* global Office */

export const selectInsertionByHost = async () => {
  let insertText;

  await Office.onReady(async (info) => {
    switch (info.host) {
      case Office.HostType.Excel: {
        insertText = insertExcelText;
        break;
      }
      case Office.HostType.PowerPoint: {
        insertText = insertPowerPointText;
        break;
      }
      case Office.HostType.Word: {
        insertText = insertWordText;
        break;
      }
      case Office.HostTyope.Outlook: {
        insertText = insertOutlookText;
      }
      default: {
        throw new Error("There is no end-to-end test for that host.");
      }
    }
  });

  return insertText;
};
