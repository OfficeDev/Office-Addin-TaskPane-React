import * as assert from "assert";
import "mocha";
import { OfficeMockObject } from "office-addin-mock";
import insertText from "../../src/word-office-document";

/* global describe, global, it, Word */

const WordMockData = {
  context: {
    document: {
      body: {
        paragraph: {
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

describe("Word", function () {
  it("Inserts text", async function () {
    const wordMock: OfficeMockObject = new OfficeMockObject(WordMockData); // Mocking the host specific namespace
    global.Word = wordMock as any;

    await insertText("Hello Word");

    wordMock.context.document.body.paragraph.load("text");
    await wordMock.context.sync();

    assert.strictEqual(wordMock.context.document.body.paragraph.text,"Hello Word");
   });
});
