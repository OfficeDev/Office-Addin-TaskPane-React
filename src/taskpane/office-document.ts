/* global Word console */
export const insertWordText = async (text: string) => {
    // Write text to the document.
    try {
        await Word.run(async (context) => {
            let body = context.document.body;
            body.insertParagraph(text, Word.InsertLocation.end);
            await context.sync();
        });
    } catch (error) {
        console.log("Error: " + error);
    }
};

/* global Excel console */
export const insertExcelText = async (text: string) => {
    // Write text to the top left cell.
    try {
        Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange("A1");
            range.values = [[text]];
            range.format.autofitColumns();
            return context.sync();
        });
    } catch (error) {
        console.log("Error: " + error);
    }
};

/* global PowerPoint console */
export const insertPowerPointText = async (text: string) => {
    // Write text to the selected slide.
    try {
        Office.context.document.setSelectedDataAsync(text, (asyncResult: Office.AsyncResult<void>) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                throw asyncResult.error.message;
            }
        });
    } catch (error) {
        console.log("Error: " + error);
    }
};

/* global Outlook console */
export const insertOutlookText = async (text: string) => {
    // Write text to the cursor point in the compose surface.
    try {
        Office.context.mailbox.item.body.setSelectedDataAsync(
            text,
            { coercionType: Office.CoercionType.Text },
            (asyncResult: Office.AsyncResult<void>) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    throw asyncResult.error.message;
                }
            }
        );
    } catch (error) {
        console.log("Error: " + error);
    }
};