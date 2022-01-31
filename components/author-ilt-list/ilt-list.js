const docx = require("docx");
const { Paragraph, TextRun } = require("docx");

const addEmptyPara = () => {
    return new Paragraph({
        children: [
            new TextRun({
                text: "",
                break: 1,
            })
        ],
        style: "normalPara"
    });
};

// regex to replace html tags with spaces, duplicate spaces into single and trim away leading and trailing
const removeTags = (string) => {
    return string.replace(/<[^>]*>/g, ' ')
        .replace(/\s{2,}/g, ' ')
        .trim();
};

const addList = module.exports = (object) => {
    const newList = [];
    const listItems = object.properties.listItems;
    // checks if whole list is empty or just the first paragraph
    if (removeTags(listItems[0].textArea) || listItems.length > 1) {
        newList.push(new Paragraph({
            children: [
                new TextRun({
                    // to be removed n production
                    text: "This is a multiple line text area displayed before the list.",
                    // text: object.properties.preTextArea,
                }),
            ],
            style: "greyedOutPara"
        }));

        listItems.forEach(listItem => {
            // replaces empty list items with empty paragraph
            if (!removeTags(listItem.textArea)) {
                newList.push(addEmptyPara());
            } else {
                newList.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: listItem.textArea.replace(/<\/?[^>]+>/gi, '')
                        }),
                    ],
                    style: "bulletPara"
                }));
            }
        });

        newList.push(new Paragraph({
            children: [
                new TextRun({
                    // to be removed n production
                    text: "This is a multiple line text area displayed after the list.",
                    // text: object.properties.postTextArea,
                }),
            ],
            style: "greyedOutPara"
        }));
    }
    else {
        newList.push(addEmptyPara());
    }
    return newList;
};


module.exports = addList;
