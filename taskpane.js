/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { base64Image } from "../../base64Image";
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("insert-image").onclick = () => clearMessage(insertImage);
    // TODO4: Assign event handler for insert-text button.
    document.getElementById("insert-text").onclick = () => clearMessage(insertText);
    // TODO6: Assign event handler for get-slide-metadata button.
    document.getElementById("get-slide-metadata").onclick = () => clearMessage(getSlideMetadata);
    // TODO8: Assign event handlers for add-slides and the four navigation buttons.
    document.getElementById("add-slides").onclick = () => tryCatch(addSlides);
    document.getElementById("go-to-first-slide").onclick = () => clearMessage(goToFirstSlide);
    document.getElementById("go-to-next-slide").onclick = () => clearMessage(goToNextSlide);
    document.getElementById("go-to-previous-slide").onclick = () => clearMessage(goToPreviousSlide);
    document.getElementById("go-to-last-slide").onclick = () => clearMessage(goToLastSlide);

    document.getElementById("saveSlideButton").addEventListener("click", () => {
      console.log("Button clicked! Triggering download.");
  
      const blob = new Blob(["This is a test file content"], { type: "text/plain" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "SlideTestFile.txt";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
  });
  
  }
});

function insertImage() {
  // Call Office.js to insert the image into the document.
  Office.context.document.setSelectedDataAsync(
    base64Image,
    {
      coercionType: Office.CoercionType.Image
    },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        setMessage("Error: " + asyncResult.error.message);
      }
    }
  );
}

// TODO5: Define the insertText function.
function insertText() {
  Office.context.document.setSelectedDataAsync("Hello World!", (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      setMessage("Error: " + asyncResult.error.message);
    }
  });
}

// TODO7: Define the getSlideMetadata function.
function getSlideMetadata() {
  Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      setMessage("Error: " + asyncResult.error.message);
    } else {
      setMessage("Metadata for selected slides: " + JSON.stringify(asyncResult.value));
    }
  });
}
// TODO9: Define the addSlides and navigation functions.
async function addSlides() {
  await PowerPoint.run(async function (context) {
    context.presentation.slides.add();
    context.presentation.slides.add();

    await context.sync();

    goToLastSlide();
    setMessage("Success: Slides added.");
  });
}

function goToFirstSlide() {
  Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      setMessage("Error: " + asyncResult.error.message);
    }
  });
}

function goToLastSlide() {
  Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      setMessage("Error: " + asyncResult.error.message);
    }
  });
}

function goToPreviousSlide() {
  Office.context.document.goToByIdAsync(Office.Index.Previous, Office.GoToType.Index, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      setMessage("Error: " + asyncResult.error.message);
    }
  });
}

function goToNextSlide() {
  Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      setMessage("Error: " + asyncResult.error.message);
    }
  });
}

async function saveSelectedSlides() {
  try {
      await PowerPoint.run(async (context) => {
          // Get the selected slides
          const presentation = context.presentation;
          const selectedSlides = presentation.getSelectedSlides();
          selectedSlides.load("items");

          await context.sync();

          if (selectedSlides.items.length === 0) {
              console.log("No slides selected.");
              return;
          }

          // You could do something like creating a new presentation from the selected slides here
          console.log(`Number of selected slides: ${selectedSlides.items.length}`);

          // For now, let's just log the number of selected slides
          // Once the selected slides are fetched, you can proceed with creating the new presentation and saving it

          // Show save file dialog (this might work only in modern browsers)
          const pickerOptions = {
              suggestedName: "NewPresentation.pptx",
              types: [{
                  description: "PowerPoint Presentation",
                  accept: { "application/vnd.openxmlformats-officedocument.presentationml.presentation": [".pptx"] }
              }]
          };

          const handle = await window.showSaveFilePicker(pickerOptions);
          const writable = await handle.createWritable();

          // This is where you'd write the new presentation's content
          // For testing, we'll just write a simple message to the file
          await writable.write("New presentation content goes here...");
          await writable.close();

          console.log("File saved successfully!");
      });
  } catch (error) {
      console.error("Error saving slides:", error);
  }
}

async function clearMessage(callback) {
  document.getElementById("message").innerText = "";
  await callback();
}

function setMessage(message) {
  document.getElementById("message").innerText = message;
}

// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
  try {
    document.getElementById("message").innerText = "";
    await callback();
  } catch (error) {
    setMessage("Error: " + error.toString());
  }
}
