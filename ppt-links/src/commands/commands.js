/* eslint-disable no-undef */
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
  window.open("https://cria.fiecon.com/", "_blank");
  event.completed();
}

function findAbbrevs(event) {
    Office.context.document.setSelectedDataAsync("Coming soon!", function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error(asyncResult.error.message);
        }
      });
      event.completed();
}




function helloWorld(event) {
  Office.context.document.setSelectedDataAsync("Hello World this is a test!", function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error(asyncResult.error.message);
    }
  });
  event.completed();
}
