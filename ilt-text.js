const docx = require("docx");
const { Paragraph, TextRun } = require("docx");

const addText = module.exports = (object) => {
    const newText = [];
    newText.push(new Paragraph({
        children: [
            new TextRun({
                text: object.displayTitle,
            }),
        ],
    }));
    return newText;
};

module.exports = addText;