/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  // Get a reference to the current message
  const urlItem = document.getElementById("tracking-pixel-url") as HTMLInputElement;
  // insert an image into the body of the message
  Office.context.mailbox.item.body.setSelectedDataAsync(`<img src="${urlItem.value}" />`, { coercionType: Office.CoercionType.Html });

  Office.context.mailbox.item.notificationMessages.addAsync("success", {
    type: "informationalMessage",
    message: "Tracking pixel inserted",
    icon: "iconid",
    persistent: false,
  });
}
