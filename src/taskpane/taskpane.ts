/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = () => {
      let errorText = document.getElementById("errorText");
      errorText.textContent = "";
      const urlItem = document.getElementById("tracking-pixel-url") as HTMLInputElement;
      const input = urlItem.value;

      if (!validateURL(input)) {
        return;
      }

      run();
    };
  }
});

function validateURL(url) {
  let regex = /^(http|https):\/\/[^ "]+$/;
  let errorText = document.getElementById("errorText");

  if (regex.test(url)) {
    return true;
  }

  errorText.textContent = "Invalid URL";
  return false;
}

export async function run() {
  // Get a reference to the current message
  const urlItem = document.getElementById("tracking-pixel-url") as HTMLInputElement;

  // check if the url is a valid url and if it is an image
  if (!urlItem.value.match(/\.(jpeg|jpg|gif|png)$/) || !urlItem.value.match(/^(http|https):\/\//)) {
    Office.context.mailbox.item.notificationMessages.addAsync("error", {
      type: "errorMessage",
      message: "Please enter a valid image url",
      icon: "iconid",
      persistent: false,
    });
    return;
  }

  // insert an image into the body of the message
  Office.context.mailbox.item.body.setSelectedDataAsync(`<img src="${urlItem.value}" />`, { coercionType: Office.CoercionType.Html });

  Office.context.mailbox.item.notificationMessages.addAsync("success", {
    type: "informationalMessage",
    message: "Tracking pixel inserted",
    icon: "iconid",
    persistent: false,
  });
}
