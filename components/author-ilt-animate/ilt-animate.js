const docx = require("docx");
const fs = require("fs");
const { Paragraph, ExternalHyperlink, TextRun } = require("docx");

const addImage = module.exports = (object) => {
    return (new Paragraph({
        children: [
            // new TextRun({
            //     text: object.displayTitle
            // }),
            new ExternalHyperlink({
                children: [
                    new TextRun({
                        text: object.displayTitle,
                        style: "Hyperlink",
                    }),
                ],
                // link: "./video.mp4", // that would work if video.mp4 is in the same folder as generated .docx document
                link: "./assets/animation.mp4",
                // link: object.properties.assetFile,
            }),
        ],
        style: "imagePara"
    }));

};

module.exports = addImage;