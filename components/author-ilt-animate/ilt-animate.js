const docx = require("docx");
const fs = require("fs");
const { Paragraph, ExternalHyperlink, TextRun } = require("docx");

const addImage = module.exports = (object) => {
    if (object.properties.adobeAnimateAsset) {

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
                    link: "./assets/animation.mp4",
                    // link: object.properties.adobeAnimateAsset,
                }),
            ],
            style: "imagePara"
        }));
    } else {
        console.log("No adobeAnimateAsset for: " + object._id);
    }

};

module.exports = addImage;