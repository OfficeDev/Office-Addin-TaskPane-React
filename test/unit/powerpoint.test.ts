import * as assert from "assert";
import "mocha";
import { OfficeMockObject } from "office-addin-mock";
import * as powerpoint from "../../src/taskpane/components/PowerPoint.App";

/* global describe, global, it */

const PowerPointMockData = {
  context: {
    document: {
      setSelectedDataAsync: function (data: string, options?) {
        this.data = data;
        this.options = options;
      },
    },
  },
  CoercionType: {
    Text: {},
  },
  onReady: async function () {},
};

describe("PowerPoint", function () {
  it("Run", async function () {
    const officeMock: OfficeMockObject = new OfficeMockObject(PowerPointMockData); // Mocking the common office-js namespace
    global.Office = officeMock as any;

    const powerpointApp = new powerpoint.default(this.props, this.context);
    await powerpointApp.click();

    assert.strictEqual(officeMock.context.document.data, "Hello World!");
  });
});
