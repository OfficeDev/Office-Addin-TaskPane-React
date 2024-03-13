import { insertText as insertTextInExcel } from "./excel";
import { insertText as insertTextInPowerPoint } from "./powerpoint";
import { insertText as insertTextInWord } from "./word";

/* global Office */

export async function insertText(host: Office.HostType, text: string) {
  switch (host) {
    case Office.HostType.Excel: {
      await insertTextInExcel(text);
      break;
    }
    case Office.HostType.PowerPoint: {
      await insertTextInPowerPoint(text);
      break;
    }
    case Office.HostType.Word: {
      await insertTextInWord(text);
      break;
    }
    default: {
      throw new Error("Don't know how to insert text when running in ${info.host}.");
    }
  }

  return insertText;
}
