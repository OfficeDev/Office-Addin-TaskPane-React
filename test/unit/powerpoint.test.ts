import * as assert from "assert";
import "mocha";
import { OfficeMockObject } from "office-addin-mock";
import { insertText } from "../../src/taskpane/powerpoint";

/* global describe, global, it */

const shapes = [];
const selectedSlide = {
  shapes: {
    addTextBox: function (text) {
      const shape = { text };
      shapes.push(shape);
    },
    items: shapes,
  },
};
const PowerPointMockData = {
  context: {
    presentation: {
      getSelectedSlides: function () {
        return {
          getItemAt: function () {
            return selectedSlide;
          },
        };
      },
    },
    slides: {
      items: [selectedSlide],
    },
  },
  onReady: async function () {},
  run: async function (callback) {
    await callback(this.context);
  },
};

describe(`PowerPoint`, function () {
  it("Inserts text", async function () {
    const officeMock = new OfficeMockObject(PowerPointMockData);
    global.PowerPoint = officeMock as any;

    await insertText("Hello PowerPoint");

    assert.strictEqual(shapes[0].text, "Hello PowerPoint");
  });
});
