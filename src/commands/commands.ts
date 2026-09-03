/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Excel, Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item?.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

/**
 * Opens taskpane.html in an Office Dialog.
 * @param event
 */
function openDialog(event: Office.AddinCommands.Event) {
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/taskpane.html",
    { height: 60, width: 40 },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error(`Dialog failed to open: ${result.error.message}`);
      }
    }
  );
  event.completed();
}

/**
 * Writes =CONTOSO.ADD(1,2) into the currently selected cell.
 * @param event
 */
function insertAddFormula(event: Office.AddinCommands.Event) {
  Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.formulas = [["=CONTOSO.ADD(1,2)"]];
    await context.sync();
  }).catch((error) => console.error(error));
  event.completed();
}

// Register the functions with Office.
Office.actions.associate("action", action);
Office.actions.associate("openDialog", openDialog);
Office.actions.associate("insertAddFormula", insertAddFormula);
