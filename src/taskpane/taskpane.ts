/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global document, Office */

/**
 * This function gets the tracking pixel from the body of the mailbox item and performs the callback function on it
 * @param callback function to perform on the tracking pixel
 */
export default function getAsyncTrackingPixel(callback: (tracking_pixel: Element) => void) {
  Office.context.mailbox.item.body.getAsync("html", {}, function (result) {
    //get the html of the body
    let html = result.value;
    //create a dummy element
    let dummy = document.createElement("div");
    //set the innerHTML of the dummy element to the html of the body
    dummy.innerHTML = html;
    //get the image with id shrunk_tracking_pixel
    let tracking_pixel = dummy.querySelector(`[title='${tracking_img_id}']`);
    //if the image exists, make shrunk-link-detected visible
    callback(tracking_pixel);
  });
}

function updateShrunkLinkDetectionMessage() {
  // check if there is an image with id shrunk_tracking_pixel in mailbox item body
  getAsyncTrackingPixel((tracking_pixel: Element) => {
    if (tracking_pixel != null) {
      document.getElementById("shrunk-link-detected").style.visibility = "visible";
      //get the url of the image, tracking_pixel
      const shrunkUrl = tracking_pixel.getAttribute("src");
      //theres shrunk-link-detected-url
      document.getElementById("shrunk-link-detected-url").textContent = shrunkUrl;
    } else {
      document.getElementById("shrunk-link-detected").style.visibility = "hidden";
    }
  });
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("insert").onclick = () => {
      let errorText = document.getElementById("errorText");
      errorText.textContent = "";

      updateShrunkLinkDetectionMessage();
      insert();
    };
    document.getElementById("remove").onclick = () => {
      remove();
    };
  }
});

const tracking_img_id = "shrunk_tracking_pixel";

export async function insert() {
  // Get a reference to the current message
  const urlItem = document.getElementById("tracking-pixel-url") as HTMLInputElement;

  // check if the url is a valid url and if it is an image
  if (!urlItem.value.match(/(https|http):\/\//)) {
    Office.context.mailbox.item.notificationMessages.replaceAsync("success", {
      type: "informationalMessage",
      message: "Please enter a valid image url",
      icon: "iconid",
      persistent: false,
    });
    return;
  }
  //check if tracking_img_id exists in body of the mailbox item
  getAsyncTrackingPixel((tracking_pixel: Element) => {
    if (tracking_pixel != null) {
      Office.context.mailbox.item.notificationMessages.replaceAsync("success", {
        type: "errorMessage",
        message: "Tracking pixel already inserted",
      });
      return;
    }
    Office.context.mailbox.item.body.setSelectedDataAsync(`<img title="${tracking_img_id}" src="${urlItem.value}" />`, {
      coercionType: Office.CoercionType.Html,
    });
    Office.context.mailbox.item.notificationMessages.replaceAsync("success", {
      type: "informationalMessage",
      message: "Tracking pixel inserted",
      persistent: false,
      icon: "iconid",
    });
  });
}

/**
 * 1. User clicks on insert
 *  a. tracking pixel already exists
 *    - error message appears
 *  b. tracking pixel does not exist
 *    - tracking pixel is inserted
 * 2. Remove button
 *   - remove if tracking pixel exists
 *   - error message appears if tracking pixel does not exist
 * Shrunk link detection --> what link did I use again ?!?!?!
 */

export async function remove() {
  // check if tracking_img_id exists in body and if it does, remove it
  updateShrunkLinkDetectionMessage();

  getAsyncTrackingPixel((tracking_pixel: Element) => {
    if (tracking_pixel == null) {
      Office.context.mailbox.item.notificationMessages.replaceAsync("success", {
        type: "errorMessage",
        message: "No tracking pixel inserted",
      });
      return;
    }
    // get rid of the image, leave everything else in the body
    Office.context.mailbox.item.body.getAsync("html", {}, function (result) {
      //get the html of the body
      let html = result.value;
      //create a dummy element
      let dummy = document.createElement("div");
      //set the innerHTML of the dummy element to the html of the body
      dummy.innerHTML = html;
      //get the image with id shrunk_tracking_pixel
      let tracking_pixel = dummy.querySelector(`[title='${tracking_img_id}']`);
      //remove the image
      tracking_pixel.remove();
      //set the body to the new html
      Office.context.mailbox.item.body.setAsync(dummy.innerHTML, { coercionType: Office.CoercionType.Html });
      // this only replaces the 'html' part of the body, not the 'text' part.
      // So any other text in the body will be preserved
    });

    Office.context.mailbox.item.notificationMessages.replaceAsync("success", {
      type: "informationalMessage",
      message: "Tracking pixel removed",
      persistent: false,
      icon: "iconid",
    });
  });
}
