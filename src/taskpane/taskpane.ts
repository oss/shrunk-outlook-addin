/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global document, Office */

/**
 * This function gets the tracking pixel from the body of the mailbox item and performs the callback function on it
 * @param callback function to perform on the tracking pixel
 */
function getAsyncTrackingPixels(callback: (tracking_pixels: NodeListOf<Element>) => void) {
  Office.context.mailbox.item.body.getAsync("html", {}, function (result) {
    let html = result.value;
    let dummy = document.createElement("div");
    dummy.innerHTML = html;
    let tracking_pixels = dummy.querySelectorAll(`img[title='${tracking_img_id}']`);
    callback(tracking_pixels);
  });
}

function getAlias(url: string) {
  // confirm that the url is from shrunk.rutgers.edu or go.rutgers.edu
  url = url.replace(/(https|http):\/\//, "");
  if (!url.match(/(shrunk|go)\.rutgers\.edu/)) {
    // get rid of the http(s)://

    // get rid of the www.
    url = url.replace("www.", "");
    return url;
  }
  let alias = url.split("/")[1];
  alias = alias.replace(".png", "");
  alias = alias.replace(".gif", "");
  return alias;
}

// function updateShrunkLinkDetectionMessage() {
//   getAsyncTrackingPixels((tracking_pixels) => {
//     if (tracking_pixels != null) {
//       document.getElementById("shrunk-link-detected").style.visibility = "visible";
//       const shrunkUrl = tracking_pixel.getAttribute("src");
//       document.getElementById("shrunk-link-detected-url").textContent = shrunkUrl;
//     } else {
//       document.getElementById("shrunk-link-detected").style.visibility = "hidden";
//     }
//   });
// }

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // loads all tracking pixels into the tracking pixel container when the extension is opened
    getAsyncTrackingPixels((tracking_pixels) => {
      let container = document.getElementById("inserted-tracking-pixels-container");
      tracking_pixels.forEach((tracking_pixel) => {
        let url = tracking_pixel.getAttribute("src");
        let newTrackingPixel = getTrackingPixelDiv(url);
        container.appendChild(newTrackingPixel);
      });
    });
    document.getElementById("insert").onclick = () => {
      let errorText = document.getElementById("errorText");
      errorText.textContent = "";
      insert();
    };
  }
});

const tracking_img_id = "shrunk_tracking_pixel";

export async function insert() {
  const urlItem = document.getElementById("tracking-pixel-url") as HTMLInputElement;

  if (!urlItem.value.match(/(https|http):\/\//)) {
    Office.context.mailbox.item.notificationMessages.replaceAsync("notify", {
      type: "informationalMessage",
      message: "Please enter a valid image url",
      icon: "iconid",
      persistent: false,
    });
    return;
  }

  //check if tracking_img_id exists in body of the mailbox item
  let error = false;
  getAsyncTrackingPixels((tracking_pixels) => {
    tracking_pixels.forEach((tracking_pixel) => {
      if (tracking_pixel.getAttribute("src") == urlItem.value) {
        Office.context.mailbox.item.notificationMessages.replaceAsync("notify", {
          type: "errorMessage",
          message: "Tracking pixel already inserted",
        });
        error = true;
        return;
      }
    });
    if (error) {
      return;
    }
    Office.context.mailbox.item.body.prependAsync(`<img title="${tracking_img_id}" src="${urlItem.value}" />`, {
      coercionType: Office.CoercionType.Html,
    });
    Office.context.mailbox.item.notificationMessages.replaceAsync("notify", {
      type: "informationalMessage",
      message: "Tracking pixel inserted",
      persistent: false,
      icon: "iconid",
    });

    let container = document.getElementById("inserted-tracking-pixels-container");
    let newTrackingPixelDiv = getTrackingPixelDiv(urlItem.value);
    container.appendChild(newTrackingPixelDiv);
  });
}

function getTrackingPixelDiv(url: string) {
  let trackingPixelDiv = document.createElement("div");
  trackingPixelDiv.title = url;

  let text = document.createElement("p");
  text.innerHTML = getAlias(url);
  trackingPixelDiv.appendChild(text);

  let removeButton = document.createElement("button");
  removeButton.textContent = "x";

  removeButton.onclick = () => {
    trackingPixelDiv.remove();
    Office.context.mailbox.item.notificationMessages.replaceAsync("notify", {
      type: "informationalMessage",
      message: "Tracking pixel removed",
      persistent: false,
      icon: "iconid",
    });

    Office.context.mailbox.item.body.getAsync("html", {}, function (result) {
      let oldHTML = result.value;
      let dummy = document.createElement("div");
      dummy.innerHTML = oldHTML;
      let allTrackingPixels = dummy.querySelectorAll(`img[title='${tracking_img_id}']`);
      allTrackingPixels.forEach((image) => {
        if (image.getAttribute("src") == trackingPixelDiv.title) {
          image.remove();
        }
      });
      Office.context.mailbox.item.body.setAsync(dummy.innerHTML, { coercionType: Office.CoercionType.Html });
    });
  };
  trackingPixelDiv.appendChild(removeButton);
  return trackingPixelDiv;
}
