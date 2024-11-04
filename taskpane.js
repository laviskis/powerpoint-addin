/* global Office, PowerPoint, PptxGenJS */

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        console.log("PowerPoint environment detected.");

        // Add an event listener with a log for debugging
        document.getElementById("saveSlideButton").onclick = () => {
            console.log("Save button clicked");
            saveSelectedSlides();
        };

        initializeApp();
    }
});


function initializeApp() {
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
console.log("PptxGenJS is available:", typeof PptxGenJS !== "undefined");

async function saveSelectedSlides() {
    try {
        const slideNumbersInput = document.getElementById("slideNumberInput").value;
        const slideNumbers = slideNumbersInput
            .split(',')
            .map((num) => parseInt(num.trim()))
            .filter((num) => !isNaN(num));

        if (slideNumbers.length === 0) {
            alert("Please enter valid slide numbers separated by commas.");
            return;
        }

        await PowerPoint.run(async (context) => {
            const presentation = context.presentation;
            const slides = presentation.slides;
            slides.load("items");

            await context.sync();

            const selectedSlides = slides.items.filter((slide, index) => slideNumbers.includes(index + 1));
            if (selectedSlides.length === 0) {
                alert("No valid slides found for the entered slide numbers.");
                return;
            }

            // Initialize new presentation with PptxGenJS
            let pptx = new PptxGenJS();

            for (let slide of selectedSlides) {
                let slideCopy = pptx.addSlide();
                // Adding the slide number or title as placeholder text
                slideCopy.addText(`Placeholder for Slide ${slide.id}`, { x: 1, y: 1, fontSize: 18 });
                console.log(`Added placeholder for Slide ${slide.id}`);
            }

            // Save the presentation as a .pptx file
            pptx.writeFile({ fileName: "SelectedSlides.pptx" });
            console.log("File saved successfully!");
        });
    } catch (error) {
        console.error("Error saving slides:", error);
    }
}

