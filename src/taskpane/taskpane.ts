/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global document, Office */

/**
 * This function gets the tracking pixel from the body of the mailbox item and performs the callback function on it
 * @param callback function to perform on the tracking pixel
 */
function getAsyncTrackingPixels(callback: (trackingPixels: NodeListOf<Element>) => void) {
  Office.context.mailbox.item.body.getAsync("html", {}, function (result) {
    let html = result.value;
    let dummy = document.createElement("div");
    dummy.innerHTML = html;
    let trackingPixels = dummy.querySelectorAll(`img[title='${TRACKING_PIXEL_TITLE}']`);
    callback(trackingPixels);
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
let insertLock = false;
let loadLock = false;
let prevChildNodes = [];

function loadTrackingPixels() {
  if (insertLock || loadLock) return;
  loadLock = true;
  getAsyncTrackingPixels((trackingPixels) => {
    let container = document.getElementById("inserted-tracking-pixels-container");
    let childNodes = [];
    trackingPixels.forEach((trackingPixel) => {
      let url = trackingPixel.getAttribute("src");
      let newTrackingPixel = getTrackingPixelDiv(url);
      childNodes.push(newTrackingPixel);
    });

    if (childNodes.length != prevChildNodes.length) {
      container.replaceChildren(...childNodes);
      prevChildNodes = childNodes;
      loadLock = false;
    } else {
      for (let i = 0; i < childNodes.length; i++) {
        if (childNodes[i].title != prevChildNodes[i].title) {
          container.replaceChildren(...childNodes);
          prevChildNodes = childNodes;
          loadLock = false;
          return;
        }
      }
      loadLock = false;
    }
  });
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // loads all tracking pixels into the tracking pixel container when the extension is opened
    loadTrackingPixels();
    document.getElementById("inserted-tracking-pixels-container").childNodes.forEach((child) => {
      prevChildNodes.push(child);
    });
    let insertButton = document.getElementById("insert");
    insertLock = false;
    insertButton.onclick = () => {
      if (insertLock) return;
      insertLock = true;
      insert();
    };
    setInterval(loadTrackingPixels, 500);
  }
});

const TRACKING_PIXEL_TITLE = "__shrunk_tracking_pixel__";

export async function insert() {
  const urlItem = document.getElementById("tracking-pixel-url") as HTMLInputElement;
  urlItem.value = urlItem.value.trim();
  let error = false;
  if (!urlItem.value.match(/(https|http):\/\//)) {
    Office.context.mailbox.item.notificationMessages.replaceAsync("notify", {
      type: "informationalMessage",
      message: "Please enter a valid image url",
      icon: "iconid",
      persistent: false,
    });
    error = true;
  }

  //check if tracking_img_id exists in body of the mailbox item

  getAsyncTrackingPixels((trackingPixels) => {
    trackingPixels.forEach((trackingPixel) => {
      if (trackingPixel.getAttribute("src") == urlItem.value) {
        Office.context.mailbox.item.notificationMessages.replaceAsync("notify", {
          type: "errorMessage",
          message: "Tracking pixel already inserted",
        });
        error = true;
        return;
      }
    });
    if (error) {
      insertLock = false; // release the lock if there is an error
      return;
    }
    Office.context.mailbox.item.body.prependAsync(
      `<img title="${TRACKING_PIXEL_TITLE}" src="${urlItem.value}" />`,
      {
        coercionType: Office.CoercionType.Html,
      },
      () => {
        // release the lock once the tracking pixel is inserted
        insertLock = false;
      }
    );
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

function setTrackingPixelBorder(borderStyle: string, src: string) {
  Office.context.mailbox.item.body.getAsync("html", {}, function (result) {
    let oldHTML = result.value;
    let dummy = document.createElement("div");
    dummy.innerHTML = oldHTML;
    let image = dummy.querySelector(`img[src='${src}']`) as HTMLImageElement;
    if (image == null) return;
    image.style.border = borderStyle;
    Office.context.mailbox.item.body.setAsync(dummy.innerHTML, { coercionType: Office.CoercionType.Html });
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

  removeButton.onclick = (event: MouseEvent) => {
    event.stopPropagation();
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
      let allTrackingPixels = dummy.querySelectorAll(`img[title='${TRACKING_PIXEL_TITLE}']`);
      allTrackingPixels.forEach((image) => {
        if (image.getAttribute("src") == trackingPixelDiv.title) {
          image.remove();
        }
      });
      Office.context.mailbox.item.body.setAsync(dummy.innerHTML, { coercionType: Office.CoercionType.Html });
    });
  };
  trackingPixelDiv.onclick = () => {
    let removeButton = trackingPixelDiv.querySelector("button");
    setTrackingPixelBorder("5px solid red", trackingPixelDiv.title);
    removeButton.disabled = true;
    trackingPixelDiv.style.pointerEvents = "none";

    setTimeout(() => {
      trackingPixelDiv.style.pointerEvents = "auto";
      trackingPixelDiv.style.border = "none";
      setTrackingPixelBorder("", trackingPixelDiv.title);
      removeButton.disabled = false;
    }, 300);
  };

  trackingPixelDiv.appendChild(removeButton);
  return trackingPixelDiv;
}
