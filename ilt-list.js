const docx = require("docx");
const { Paragraph, TextRun } = require("docx");

const addList = module.exports = (object) => {
    // console.log(object);
    const newList = [];
    const listItems = object.properties.listItems;
    if (listItems[0].textArea !== "<p></p>") {
        newList.push(new Paragraph({
            children: [
                new TextRun({
                    text: "This is a multiple line text area displayed before the list.",
                    color: "#808080",
                }),
            ]
        }));
    }

    listItems.forEach(listItem => {
        newList.push(new Paragraph({
            children: [
                new TextRun({
                    text: listItem.textArea.replace(/<\/?[^>]+>/gi, '')
                }),
            ],
            bullet: {
                level: 0
            },
            style: "normalPara"
        }));

    });

    if (listItems[0].textArea !== "<p></p>") {
        newList.push(new Paragraph({
            children: [
                new TextRun({
                    text: "This is a multiple line text area displayed after the list.",
                    color: "#808080"
                }),
            ]
        }));
    }

    return newList;
};

module.exports = addList;
