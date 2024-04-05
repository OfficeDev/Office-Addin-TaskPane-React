/* global Office console */

export async function insertText(text: string) {
  // Write text to the cursor point in the compose surface.
  try {
    Office.context.mailbox.item?.body.setSelectedDataAsync(
      text,
      { coercionType: Office.CoercionType.Text },
      (asyncResult: Office.AsyncResult<void>) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          throw asyncResult.error.message;
        }
      }
    );
  } catch (error) {
    console.log("Error: " + error);
  }
}
