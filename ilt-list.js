const docx = require("docx");
const { Paragraph, TextRun } = require("docx");
const addList = module.exports = (object) => {
    // console.log(object);
    const newList = [];
    object.properties.listItems.forEach(listItem => { //it assumes all lists have bullet points - to FIX!!!
        newList.push(new Paragraph({
            children: [
                // new TextRun({
                //     text: "This is a multiple line text area displayed before the list."
                // }),
                new TextRun({
                    text: listItem.textArea.replace(/<\/?[^>]+>/gi, '')
                }),
                // new TextRun({
                //     text: "This is a multiple line text area displayed after the list."
                // })
            ],
            bullet: {
                level: 0
            },
            // style: "normalPara"
        }));

    });
    // console.log(newList);

    return newList;
};

module.exports = addList;
