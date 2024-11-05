Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        console.log("PowerPoint environment detected.");
        document.getElementById("saveSlideButton").onclick = createNewPresentationWithSelectedSlides;
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

async function createNewPresentationWithSelectedSlides() {
    const slideNumbersInput = document.getElementById("slideNumberInput").value;
    const slideNumbers = slideNumbersInput
        .split(',')
        .map((num) => parseInt(num.trim()))
        .filter((num) => !isNaN(num));

    if (slideNumbers.length === 0) {
        alert("Please enter valid slide numbers separated by commas.");
        return;
    }

    try {
        await PowerPoint.run(async (context) => {
            const presentation = context.presentation;
            const slides = presentation.slides;
            slides.load("items");

            await context.sync();

            const totalSlides = slides.items.length;
            const validSlideNumbers = slideNumbers.filter(num => num > 0 && num <= totalSlides);

            if (validSlideNumbers.length === 0) {
                alert("No valid slides found for the entered slide numbers.");
                return;
            }

            // Initialize a new PptxGenJS presentation
            let pptx = new PptxGenJS();

            // Add each selected slide as a placeholder in the new presentation
            validSlideNumbers.forEach((slideNum) => {
                let slideCopy = pptx.addSlide();
                slideCopy.addText(`Placeholder for Slide #${slideNum}`, { x: 1, y: 1, fontSize: 18 });
            });

            // Save the new presentation as a .pptx file and trigger download
            pptx.writeFile({ fileName: "SelectedSlidesPresentation.pptx" }).then((filePath) => {
                console.log("Presentation created successfully:", filePath);
                openOutlookWithAttachment(filePath); // Trigger Outlook email with attachment
            });
        });
    } catch (error) {
        console.error("Error creating presentation with selected slides:", error);
    }
}

function openOutlookWithAttachment() {
    const subject = "Slides from PowerPoint Presentation";
    const body = "Please find the selected slides from the PowerPoint presentation attached.";
    const mailtoLink = `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;

    // Open the default email client (e.g., Outlook) with a prefilled subject and body
    window.location.href = mailtoLink;

    // Inform the user to attach the file manually
    alert("A new email has been opened in your default email client. Please attach the downloaded .pptx file manually.");
}
