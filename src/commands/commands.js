/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

async function toggleProtection(args) {
  try {
    await Excel.run(async (context) => {
      // TODO1: Queue commands to reverse the protection status of the current worksheet.
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // TODO2: Queue command to load the sheet's "protection.protected" property from
      //        the document and re-synchronize the document and task pane.
      sheet.load("protection/protected");
      await context.sync();

      if (sheet.protection.protected) {
        sheet.protection.unprotect();
      } else {
        sheet.protection.protect();
      }

      await context.sync();
    });
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }

  args.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);
Office.actions.associate("toggleProtection", toggleProtection);
