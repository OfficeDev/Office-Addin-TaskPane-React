import * as assert from "assert";
import "mocha";
import { OfficeMockObject } from "office-addin-mock";
import insertText from "../../src/excel-office-document";

/* global describe, global, it */

const ExcelMockData = {
  context: {
    workbook: {
      worksheets: {
        range: {
          values: [[" "]],
          format: {
            autoFitColumns: function () { }
          },
        },
        getRange: function () {
          return this.range;
        },
        getActiveWorksheet: function () {
          return this;
        }
      }
    }
  },
  run: async function (callback) {
    await callback(this.context);
  },
};

describe("Excel", function () {
  it("Inserts text", async function () {
    const excelMock: OfficeMockObject = new OfficeMockObject(ExcelMockData); // Mocking the host specific namespace
    global.Excel = excelMock as any;

    await insertText("Hello Excel");

    excelMock.context.workbook.worksheets.range.load("values");
    await excelMock.context.sync();

   assert.strictEqual(excelMock.context.workbook.worksheets.range.values[0][0], "Hello Excel");
  });
});
