/* global Office */

const insertText = async (text: string) => {
  // Write text to the selected slide.
  return Office.context.document.setSelectedDataAsync(
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
