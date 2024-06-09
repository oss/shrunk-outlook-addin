/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global document, Office */

const ONLY_ALLOW_RUTGERS_LINKS = false;

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
let highlightLock = false;
let prevChildNodes = [];

function loadTrackingPixels() {
  if (insertLock || loadLock || highlightLock) return;
  loadLock = true;
  getAsyncTrackingPixels((trackingPixels) => {
    let container = document.getElementById("inserted-tracking-pixels-container");
    let childNodes = [];

    if (trackingPixels.length == 0) {
      document.getElementById("no-inserted-detected").style.display = "block";
      document.getElementById("inserted-detection-instruction").style.display = "none";
    } else {
      document.getElementById("no-inserted-detected").style.display = "none";
      document.getElementById("inserted-detection-instruction").style.display = "block";
    }

    trackingPixels.forEach((trackingPixel) => {
      let url = trackingPixel.getAttribute("src");
      let newTrackingPixel = createTrackingPixelDiv(url);
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

function setupInstructionDropdown() {
  let dropdownButton = document.getElementById("dropdown");
  dropdownButton.onclick = () => {
    let dropdownDiv = document.getElementById("instruction-list");
    dropdownDiv.style.maxHeight = dropdownDiv.style.maxHeight == "500px" ? "0" : "500px";
    dropdownButton.style.transform =
      dropdownButton.style.transform == "rotate(180deg) translateY(6px)"
        ? "rotate(0deg)"
        : "rotate(180deg) translateY(6px)";
  };
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

    let textItem = document.getElementById("tracking-pixel-url") as HTMLInputElement;
    textItem.addEventListener("keyup", function (event) {
      if (event.key === "Enter") {
        event.preventDefault();
        insertButton.click();
      }
    });

    setInterval(loadTrackingPixels, 500);

    setupInstructionDropdown();
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
  if (ONLY_ALLOW_RUTGERS_LINKS && !urlItem.value.match(/^https?:\/\/(shrunk|go)\.rutgers\.edu\/.*\.(png|gif)$/)) {
    Office.context.mailbox.item.notificationMessages.replaceAsync("notify", {
      type: "errorMessage",
      message: "Only go or shrunk links are allowed",
    });
    error = true;
  }
  if (error) {
    insertLock = false; // release the lock if there is an error
    return;
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
    let newTrackingPixelDiv = createTrackingPixelDiv(urlItem.value);

    container.appendChild(newTrackingPixelDiv);
  });
}

function setTrackingPixelBorder(borderStyle: string, src: string, callback: () => void) {
  Office.context.mailbox.item.body.getAsync("html", {}, function (result) {
    let oldHTML = result.value;
    let dummy = document.createElement("div");
    dummy.innerHTML = oldHTML;
    let image = dummy.querySelector(`img[src='${src}']`) as HTMLImageElement;
    if (image == null) return;
    image.style.border = borderStyle;
    Office.context.mailbox.item.body.setAsync(dummy.innerHTML, { coercionType: Office.CoercionType.Html }, callback);
  });
}

function createTrackingPixelDiv(url: string) {
  let trackingPixelDiv = document.createElement("div");
  trackingPixelDiv.title = url;

  let text = document.createElement("p");
  text.innerHTML = getAlias(url);
  trackingPixelDiv.appendChild(text);

  let removeButton = document.createElement("button");
  // add the close_svg to the button
  let img = document.createElement("img");
  /*
    IMPORTANT: 
    
    Due to time constraints, we are using an actual live link to download the delete_svg file.
    In normal circumstances, we would use a relative path towards the file. However,
    there are some issues when the distribution is created upon release. 

    If there are issues with the delete icon, please investigate how the path towards assets
    change and create the necessary modifications to this line.

    Please change this in the future.
  */
  img.src = "https://shrunk.rutgers.edu/outlook/assets/dev/assets/delete_svg.svg"; // the svg is located in ../../assets/delete_svg.svg
  removeButton.appendChild(img);

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
    if (highlightLock) return;
    let removeButton = trackingPixelDiv.querySelector("button");
    setTrackingPixelBorder("5px solid red", trackingPixelDiv.title, () => { });
    removeButton.disabled = true;
    trackingPixelDiv.style.pointerEvents = "none";
    highlightLock = true;
    setTimeout(() => {
      trackingPixelDiv.style.pointerEvents = "auto";
      trackingPixelDiv.style.border = "none";
      setTrackingPixelBorder("", trackingPixelDiv.title, () => {
        highlightLock = false;
      });
      removeButton.disabled = false;
    }, 200);
  };

  trackingPixelDiv.appendChild(removeButton);
  return trackingPixelDiv;
}
