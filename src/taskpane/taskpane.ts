import { insertText as insertTextInExcel } from "./excel";
import { insertText as insertTextInOneNote } from "./onenote";
import { insertText as insertTextInOutlook } from "./outlook";
import { insertText as insertTextInPowerPoint } from "./powerpoint";
import { insertText as insertTextInProject } from "./project";
import { insertText as insertTextInWord } from "./word";

/* global Office */

export async function insertText(text: string) {
  Office.onReady(async (info) => {
    switch (info.host) {
      case Office.HostType.Excel:
        await insertTextInExcel(text);
        break;
      case Office.HostType.OneNote:
        await insertTextInOneNote(text);
        break;
      case Office.HostType.Outlook:
        await insertTextInOutlook(text);
        break;
      case Office.HostType.Project:
        await insertTextInProject(text);
        break;
      case Office.HostType.PowerPoint:
        await insertTextInPowerPoint(text);
        break;
      case Office.HostType.Word:
        await insertTextInWord(text);
        break;
      default: {
        throw new Error("Don't know how to insert text when running in ${info.host}.");
      }
    }
  });
}
