const docx = require("docx");
const fs = require("fs");
const { Paragraph, TextRun, ImageRun } = require("docx");

const addDragImage = module.exports = (object) => {
    const newDragImage = [];
    if (object.properties.imageField) {
        newDragImage.push(new Paragraph({
            children: [
                new TextRun({
                    text: object.properties.question.replace(/<\/?[^>]+>/gi, ''),
                    bold: true
                }),
            ],
        }));
        newDragImage.push(new Paragraph({
            children: [
                new ImageRun({
                    data: fs.readFileSync("./assets/question.jpg"),
                    // data: fs.readFileSync(object.properties.imageField),
                    transformation: {
                        width: 600,
                        height: 350,
                    },
                })
            ],
            style: "imagePara"
        })
        );
        return newDragImage;
    } else {
        console.log("No asses for: " + object._id);
    }
};

module.exports = addDragImage;