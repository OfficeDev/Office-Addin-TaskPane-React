import * as assert from "assert";
import "mocha";
import { OfficeMockObject } from "office-addin-mock";
import insertText from "../../src/taskpane/powerpoint-office-document";

/* global describe, global, it */

const PowerPointMockData = {
  context: {
    document: {
      setSelectedDataAsync: function (data: string, callback?) {
        this.data = data;
        this.callback = callback;
      },
    },
  },
  CoercionType: {
    Text: {},
  },
  onReady: async function () {},
};

describe(`PowerPoint`, function () {
  it("Inserts text", async function () {
    const officeMock = new OfficeMockObject(PowerPointMockData);
    global.Office = officeMock as any;

    await insertText("Hello PowerPoint");

    assert.strictEqual(officeMock.context.document.data, "Hello PowerPoint");
  });
});
