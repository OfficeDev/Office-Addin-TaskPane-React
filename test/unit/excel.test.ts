import * as assert from "assert";
import "mocha";
import { OfficeMockObject } from "office-addin-mock";
import * as excel from "../../src/taskpane/components/excel.App";

/* global describe, global, it */

const ExcelMockData = {
  context: {
    workbook: {
      range: {
        address: "G4",
        format: {
          fill: {},
        },
      },
      getSelectedRange: function () {
        return this.range;
      },
    },
  },
  run: async function (callback) {
    await callback(this.context);
  },
};

const OfficeMockData = {
  onReady: async function () {},
};

describe("Excel", function () {
  it("Run", async function () {
    const excelMock: OfficeMockObject = new OfficeMockObject(ExcelMockData); // Mocking the host specific namespace
    global.Excel = excelMock as any;
    global.Office = new OfficeMockObject(OfficeMockData) as any; // Mocking the common office-js namespace

    const excelApp = new excel.default(this.props, this.context);
    await excelApp.click();

    assert.strictEqual(excelMock.context.workbook.range.format.fill.color, "yellow");
  });
});
