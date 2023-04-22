import * as assert from "assert";
import "mocha";
import { OfficeMockObject } from "office-addin-mock";
import * as word from "../../src/taskpane/components/Word.App";

/* global describe, global, it, Word */

const WordMockData = {
  context: {
    document: {
      body: {
        paragraph: {
          font: {},
          text: "",
        },
        insertParagraph: function (paragraphText: string, insertLocation: Word.InsertLocation): Word.Paragraph {
          this.paragraph.text = paragraphText;
          this.paragraph.insertLocation = insertLocation;
          return this.paragraph;
        },
      },
    },
  },
  InsertLocation: {
    end: "End",
  },
  run: async function (callback) {
    await callback(this.context);
  },
};

const OfficeMockData = {
  onReady: async function () { },
};

describe("Word", function () {
  it("Run", async function () {
    const wordMock: OfficeMockObject = new OfficeMockObject(WordMockData); // Mocking the host specific namespace
    global.Word = wordMock as any;
    global.Office = new OfficeMockObject(OfficeMockData) as any; // Mocking the common office-js namespace

    const wordApp = new word.default(this.props, this.context);
    await wordApp.click();

    assert.strictEqual(wordMock.context.document.body.paragraph.font.color, "blue");
  });
});
