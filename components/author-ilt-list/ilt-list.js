const docx = require("docx");
const { Paragraph, TextRun } = require("docx");

const addEmptyPara = () => {
    console.log("added empty");
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
    // if (listItems[0].textArea.replace(/\s+/g, '') !== "<p></p>") {
    if (removeTags(listItems[0].textArea)) {
        // console.log("first non empty");
        newList.push(new Paragraph({
            children: [
                new TextRun({
                    text: "This is a multiple line text area displayed before the list.",
                    color: "#808080",
                }),
            ],
            style: "greyedOutPara"
        }));

        listItems.forEach(listItem => {
            if (!removeTags(listItem.textArea)) {
                console.log("not first empty");
                console.log(listItem.textArea);
                addEmptyPara();
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
                    text: "This is a multiple line text area displayed after the list.",
                    color: "#808080"
                }),
            ],
            style: "greyedOutPara"

        }));

    }
    else {
        newList.push(addEmptyPara());
        // console.log("empty para");
    }
    return newList;

};


module.exports = addList;
