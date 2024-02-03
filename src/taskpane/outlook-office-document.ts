/* global Office */

const insertText = async (text: string) => {
  // Write text to the cursor point in the compose surface.
  return Office.context.mailbox.item.body.setSelectedDataAsync(
    text,
    { coercionType: Office.CoercionType.Text },
    (asyncResult: Office.AsyncResult<void>) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        throw asyncResult.error.message;
      }
    }
  );
};

export default insertText;