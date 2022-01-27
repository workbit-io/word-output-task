const docx = require("docx");
const { Paragraph, TextRun } = require("docx");

const addText = module.exports = (object) => {
    const newMcq = [];
    newMcq.push(new Paragraph({
        children: [
            new TextRun({
                text: object.properties.question.replace(/<\/?[^>]+>/gi, ''),
            }),
        ],
    }));
    object.properties.answers.forEach(answer => {
        newMcq.push(new Paragraph({
            children: [
                new TextRun({
                    text: answer.answerText.replace(/<\/?[^>]+>/gi, '')
                }),
            ],
            bullet: {
                level: 0
            },
            style: "normalPara"
        }));

    });
    return newMcq;
};

module.exports = addText;