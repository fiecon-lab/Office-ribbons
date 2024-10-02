/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */



Office.onReady(() => {
    
  // Register the function with Office.
  Office.actions.associate("openCria", openCria);
  Office.actions.associate("findAbbrevs", findAbbrevs);
});


function openCria(event) {
  window.open('https://cria.fiecon.com/', '_blank');
  event.completed();
}


function findAbbrevs(event) {

  console.log("Hello world");

  // const message = {
  //   type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
  //   message: "Performed action.",
  //   icon: "Icon.80x80",
  //   persistent: true,
  // };

  // // Show a notification message.
  // Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);
 
  // event.completed();
}





function helloWorld(event) {
  Office.context.document.setSelectedDataAsync(
      "Hello World!",
      function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              console.error(asyncResult.error.message);
          }
      }
  );
  event.completed();
}

