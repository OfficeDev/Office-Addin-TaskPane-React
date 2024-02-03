/* global Office console */

const insertText = async (text: string) => {
  // Write text to the task notes field.
  return Office.context.document.getSelectedTaskAsync((result: Office.AsyncResult<string>) => {
    // Get the GUID of the selected task
    let taskGuid: string;
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      taskGuid = result.value;

      // Set the specified fields for the selected task.
      const targetFields: Office.ProjectTaskFields[] = [Office.ProjectTaskFields.Name, Office.ProjectTaskFields.Notes];
      const fieldValues: string[] = ["New task name", text];

      // Set the field value. If the call is successful, set the next field.
      for (let index = 0; index < targetFields.length; index++) {
        Office.context.document.setTaskFieldAsync(
          taskGuid,
          targetFields[index],
          fieldValues[index],
          (result: Office.AsyncResult<void>) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              index++;
            } else {
              console.log(result.error);
            }
          }
        );
      }
    } else {
      throw result.error.message;
    }
  });
};

export default insertText;
