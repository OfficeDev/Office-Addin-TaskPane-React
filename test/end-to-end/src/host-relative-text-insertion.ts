import { default as insertExcelText} from "../../../src/excel-office-document";
import { default as insertPowerPointText} from "../../../src/powerpoint-office-document";
import { default as insertWordText} from "../../../src/word-office-document";

/* global Office */

let insertText;

Office.onReady(async (info) => {
  switch (info.host) {
    case Office.HostType.Excel: {
      insertText = insertExcelText;
    }
    case Office.HostType.PowerPoint: {
      insertText = insertPowerPointText;
    }
    case Office.HostType.Word: {
      insertText = insertWordText;
    }
    default: {
      throw new Error("There is no end-to-end test for that host.");
    }
  }
});

export const selectInsertionByHost = () => {
  return insertText;
}