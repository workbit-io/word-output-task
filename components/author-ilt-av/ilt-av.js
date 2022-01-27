const docx = require("docx");
const fs = require("fs");
const { Paragraph, ExternalHyperlink, TextRun } = require("docx");

const addImage = module.exports = (object) => {
    return (new Paragraph({
        children: [
            new ExternalHyperlink({
                children: [
                    new TextRun({
                        text: "External link to video",
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

};

module.exports = addImage;