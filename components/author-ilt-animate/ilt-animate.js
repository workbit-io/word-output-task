const docx = require("docx");
const fs = require("fs");
const { Paragraph, ExternalHyperlink, TextRun } = require("docx");

const addImage = module.exports = (object) => {
    return (new Paragraph({
        children: [
            new TextRun({
                text: object.displayTitle
            })
        ],
        style: "imagePara"
    }));

};

module.exports = addImage;