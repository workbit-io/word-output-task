const docx = require("docx");
const { Paragraph, TextRun } = require("docx");

const adds1000d = module.exports = (object) => {
    const newText = [];
    const title = object.displayTitle;
    newText.push(new Paragraph({
        children: [
            new TextRun({
                text: title,
            }),
            new TextRun({
                text: object.properties.preTextArea.replace(/<\/?[^>]+>/gi, ''),
                break: 1,
            }),
            new TextRun({
                text: object.properties.textAreaField.replace(/<\/?[^>]+>/gi, ''),
                break: 1,
            }),
            new TextRun({
                text: object.properties.postTextArea.replace(/<\/?[^>]+>/gi, ''),
                break: 1,
            }),
        ],
        style: object.title,
    }));


    return newText;
};

module.exports = adds1000d;