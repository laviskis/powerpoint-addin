/* global Office, PowerPoint */

// Function to initialize the app
function initializeApp() {
    console.log("Initializing app...");

    const sideloadMsg = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");

    if (sideloadMsg && appBody) {
        sideloadMsg.style.display = "none";
        appBody.style.display = "flex";
        console.log("App body displayed successfully.");
    } else {
        console.error("sideload-msg or app-body element not found in the DOM");
    }
}

// Office initialization
Office.onReady((info) => {
    try {
        if (info.host === Office.HostType.PowerPoint) {
            console.log("PowerPoint environment detected.");
            initializeApp();

            document.getElementById("insert-image").onclick = () => clearMessage(insertImage);
            document.getElementById("insert-text").onclick = () => clearMessage(insertText);
            document.getElementById("get-slide-metadata").onclick = () => clearMessage(getSlideMetadata);
            document.getElementById("add-slides").onclick = () => tryCatch(addSlides);
            document.getElementById("go-to-first-slide").onclick = () => clearMessage(goToFirstSlide);
            document.getElementById("go-to-next-slide").onclick = () => clearMessage(goToNextSlide);
            document.getElementById("go-to-previous-slide").onclick = () => clearMessage(goToPreviousSlide);
            document.getElementById("go-to-last-slide").onclick = () => clearMessage(goToLastSlide);
            document.getElementById("saveSlideButton").onclick = () => triggerDownload();
        }
    } catch (error) {
        console.error("Error during Office.onReady:", error);
    }
});

function insertImage() {
    Office.context.document.setSelectedDataAsync(
        base64Image,
        { coercionType: Office.CoercionType.Image },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                setMessage("Error: " + asyncResult.error.message);
            }
        }
    );
}

function insertText() {
    Office.context.document.setSelectedDataAsync("Hello World!", (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            setMessage("Error: " + asyncResult.error.message);
        }
    });
}

function getSlideMetadata() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            setMessage("Error: " + asyncResult.error.message);
        } else {
            setMessage("Metadata for selected slides: " + JSON.stringify(asyncResult.value));
        }
    });
}

async function addSlides() {
    await PowerPoint.run(async (pptContext) => {
        pptContext.presentation.slides.add();
        pptContext.presentation.slides.add();
        await pptContext.sync();
        goToLastSlide();
        setMessage("Success: Slides added.");
    });
}

function goToFirstSlide() {
    Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index, handleAsyncResult);
}

function goToLastSlide() {
    Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index, handleAsyncResult);
}

function goToPreviousSlide() {
    Office.context.document.goToByIdAsync(Office.Index.Previous, Office.GoToType.Index, handleAsyncResult);
}

function goToNextSlide() {
    Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index, handleAsyncResult);
}

function handleAsyncResult(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        setMessage("Error: " + asyncResult.error.message);
    }
}

function triggerDownload() {
    const blob = new Blob(["This is a test file content"], { type: "text/plain" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "SlideTestFile.txt";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function clearMessage(callback) {
    document.getElementById("message").innerText = "";
    callback();
}

function setMessage(message) {
    document.getElementById("message").innerText = message;
}

async function tryCatch(callback) {
    try {
        document.getElementById("message").innerText = "";
        await callback();
    } catch (error) {
        setMessage("Error: " + error.toString());
    }
}
