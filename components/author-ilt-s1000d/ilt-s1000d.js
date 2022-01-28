const docx = require("docx");
const { Paragraph, TextRun, BorderStyle } = require("docx");

const adds1000d = module.exports = (object) => {
    console.log("called");
    const newText = [];
    const title = object.displayTitle;
    newText.push(new Paragraph({
        children: [
            new TextRun({
                text: title,
                break: 1,
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
        style: "normalPara",
        border: {
            top: {
                color: "#FF0000",
                space: 1,
                style: BorderStyle.DASHED,
                size: 15,
            },
            bottom: {
                color: "#FF0000",
                space: 1,
                style: BorderStyle.DASHED,
                size: 15,
            },
            // left: {
            //     color: "#FF0000",
            //     space: 10,
            //     style: BorderStyle.DASHED,
            //     size: 15,
            // },
            // right: {
            //     color: "#FF0000",
            //     space: 10,
            //     style: BorderStyle.DASHED,
            //     size: 15,
            // },
        },
    }));


    return newText;
};

module.exports = adds1000d;