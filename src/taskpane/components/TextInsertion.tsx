import * as React from "react";
import { useState } from "react";
import { Button, Field, Textarea, tokens, makeStyles } from "@fluentui/react-components";
import { insertWordText, insertExcelText, insertPowerPointText, insertOutlookText } from "../office-document";

const useStyles = makeStyles({
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    marginTop: "20px",
    marginBottom: "10px",
  },
  textPromptAndInsertion: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
  textAreaField: {
    marginLeft: "20px",
    marginTop: "30px",
    marginBottom: "20px",
    marginRight: "20px",
    maxWidth: "50%",
  },
});

const TextInsertion: React.FC = () => {
  const [text, setText] = useState<string>("Some text.");

  const handleTextInsertion = async () => {
    const insertText = await selectInsertionByHost();
    await insertText(text);
  };

  const handleTextChange = async (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    setText(event.target.value);
  };

  const styles = useStyles();

  const selectInsertionByHost = async () => {
    let insertText;

    await Office.onReady(async (info) => {
      switch (info.host) {
        case Office.HostType.Excel: {
          insertText = insertExcelText;
          break;
        }
        case Office.HostType.PowerPoint: {
          insertText = insertPowerPointText;
          break;
        }
        case Office.HostType.Word: {
          insertText = insertWordText;
          break;
        }
        case Office.HostType.Outlook: {
          insertText = insertOutlookText;
        }
        default: {
          throw new Error("There is no end-to-end test for that host.");
        }
      }
    });

    return insertText;
  };

  return (
    <div className={styles.textPromptAndInsertion}>
      <Field className={styles.textAreaField} size="large" label="Enter text to be inserted into the document.">
        <Textarea size="large" value={text} onChange={handleTextChange} />
      </Field>
      <Field className={styles.instructions}>Click the button to insert text.</Field>
      <Button appearance="primary" disabled={false} size="large" onClick={handleTextInsertion}>
        Insert text
      </Button>
    </div>
  );
};

export default TextInsertion;
