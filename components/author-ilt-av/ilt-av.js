const docx = require("docx");
const fs = require("fs");
const { Paragraph, ExternalHyperlink, TextRun } = require("docx");

const addAV = module.exports = (object) => {
    if (object.properties.assetFile) {
        return (new Paragraph({
            children: [
                new ExternalHyperlink({
                    children: [
                        new TextRun({
                            text: object.displayTitle,
                            style: "Hyperlink",
                        }),
                    ],
                    // link: "./video.mp4", // that would work if video.mp4 is in the same folder as generated .docx document
                    // link: "./assets/video.mp4",
                    link: object.properties.assetFile,
                }),
            ],
            style: "imagePara"
        }));
    } else {
        console.log("No asset file for: " + object._id);
    }

};

module.exports = addAV;