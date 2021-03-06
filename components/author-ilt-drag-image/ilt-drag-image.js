const docx = require("docx");
const fs = require("fs");
const { Paragraph, TextRun, ImageRun } = require("docx");

const addDragImage = module.exports = (object) => {
    const newDragImage = [];
    newDragImage.push(new Paragraph({
        children: [
            new TextRun({
                text: object.properties.question.replace(/<\/?[^>]+>/gi, ''),
            }),
        ],
    }));
    if (object.properties.imageField) {
        newDragImage.push(new Paragraph({
            children: [
                new ImageRun({
                    data: fs.readFileSync("./assets/question.jpg"),
                    transformation: {
                        width: 600,
                        height: 350,
                    },
                })
            ]
        })
        );
    }
    return newDragImage;
};

module.exports = addDragImage;