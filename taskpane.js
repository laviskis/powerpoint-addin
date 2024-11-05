Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        console.log("PowerPoint environment detected.");
        document.getElementById("saveSlideButton").onclick = filterSlidesInPresentation;
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

async function filterSlidesInPresentation() {
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

            // Create a list of slide indices to delete (those not in the selected list)
            const slidesToDelete = [];
            for (let i = 0; i < totalSlides; i++) {
                if (!validSlideNumbers.includes(i + 1)) {
                    slidesToDelete.push(slides.items[i]);
                }
            }

            // Delete slides that are not in the selected list
            slidesToDelete.forEach((slide) => {
                slide.delete();
            });

            await context.sync();

            console.log("Unwanted slides removed successfully!");
            alert("Unwanted slides have been removed from the presentation.");
        });
    } catch (error) {
        console.error("Error filtering slides in the presentation:", error);
        alert("An error occurred while filtering slides.");
    }
}
