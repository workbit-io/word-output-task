const docx = require("docx");
const { Paragraph, TextRun } = require("docx");

const addText = module.exports = (object) => {
    // console.log(object.displayTitle);
    if (object.displayTitle) {
        return (new Paragraph({
            children: [
                new TextRun({
                    text: object.displayTitle,
                }),
            ],
        })
        );
    }
};

module.exports = addText;