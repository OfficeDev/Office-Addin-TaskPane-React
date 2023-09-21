import { default as insertExcelText } from "./taskpane/excel-office-document";
import { default as insertPowerPointText } from "./taskpane/powerpoint-office-document";
import { default as insertWordText } from "./taskpane/word-office-document";

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
      default: {
        throw new Error("There is no end-to-end test for that host.");
      }
    }
  });

  return insertText;
};
