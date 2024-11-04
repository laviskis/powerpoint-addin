Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        console.log("PowerPoint environment detected.");
        document.getElementById("saveSlideButton").onclick = duplicateAndTrimPresentation;
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

async function duplicateAndTrimPresentation() {
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

            const totalSlides = slides.items.length;

            // Validate slide numbers to be within range
            const validSlideNumbers = slideNumbers.filter(num => num > 0 && num <= totalSlides);
            if (validSlideNumbers.length === 0) {
                alert("No valid slides found for the entered slide numbers.");
                return;
            }

            // Start with a copy of the entire presentation
            let pptx = new PptxGenJS();
            let copiedSlides = [];

            // Create a new slide in PptxGenJS for each slide in the original presentation
            for (let i = 0; i < totalSlides; i++) {
                let slideCopy = pptx.addSlide();
                slideCopy.addText(`Placeholder for Slide #${i + 1}`, { x: 1, y: 1, fontSize: 18 });
                copiedSlides.push(slideCopy);
            }

            // Remove slides that are not in the list of selected slides
            for (let i = 0; i < copiedSlides.length; i++) {
                if (!validSlideNumbers.includes(i + 1)) {
                    pptx.slides.splice(i, 1);
                }
            }

            // Save the trimmed presentation
            pptx.writeFile({ fileName: "TrimmedPresentation.pptx" });
            console.log("File saved successfully!");
        });
    } catch (error) {
        console.error("Error duplicating and trimming presentation:", error);
    }
}
