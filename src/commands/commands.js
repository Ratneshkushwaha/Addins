/* eslint-disable no-debugger */
/* eslint-disable prettier/prettier */
/* eslint-disable no-undef */
/* eslint-disable office-addins/no-office-initialize */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
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

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;
// The initialize function is required for all apps.
Office.initialize = function () {
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    console.log(subject);
    const masterCategoriesToAdd = [
      {
        displayName: "HelloCategory",
        color: Office.MailboxEnums.CategoryColor.Preset0
      }
    ];   
    Office.context.mailbox.Folder.addAsync(masterCategoriesToAdd, function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added Folder to master list");
      } else {
        console.log("test Folder.addAsync call failed with error: " + asyncResult.error.message);
      }
    });
    Office.context.mailbox.masterCategories.addAsync(masterCategoriesToAdd, function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories to master list");
      } else {
        console.log("masterCategories.addAsync call failed with error: " + asyncResult.error.message);
      }
    });
    
    // Office.context.mailbox.masterCategories.getAsync(function(asyncResult) {
    //   if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    //     const categories = asyncResult.value;
    //     if (categories && categories.length > 0) {
    //       console.log("Master categories:");
    //       console.log(JSON.stringify(categories));
    //     } else {
    //       console.log("There are no categories in the master list.");
    //     }
    //   } else {
    //     console.error(asyncResult.error);
    //   }
    // });
};
