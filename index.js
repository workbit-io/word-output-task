const docx = require("docx");
const fs = require("fs");
// const data = require("./wlodek_json_word.json");
const data = require("./telebrief.json");
const { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel, TableOfContents, Header, Footer, TextWrappingType, TextWrappingSide, PageNumber, AlignmentType, BorderStyle } = require("docx");

let doc;
const contents = [];
let previousHeading = 1;

// keep the track of heading numbering
let nextHeading1Num = 0;
let nextHeading2Num = 0;
let nextHeading3Num = 0;

const paragraphStyles = {
    heading1: {
        id: "WorkbitHeading1",
        name: "WorkbitHeading1",
        basedOn: "Heading1",
        next: "Heading1",
        quickFormat: true,
        run: {
            font: "Arial",
            size: 28,
            bold: true,
            color: "#FF0000",
        }
    }
};

// extracts course from the json file
const createWordOutput = () => {
    const course = data.content[0];
    getHeadings(course);
};

const generateHeading1 = (article) => {
    const sessionIntroductionTitle = article.children[0].children[0];
    // setting next HEADING numbering to +1
    // resetting HEADING 2 and HEADING 3 as with every article/Section Introduction element numbering starts from beginning
    nextHeading1Num++;
    nextHeading2Num = 0;
    nextHeading3Num = 0;
    // console.log(nextHeading1Num);
    // console.log(sessionIntroductionTitle[0].displayTitle);
    createHeading(sessionIntroductionTitle.displayTitle, nextHeading1Num, 1);
};

const generateHeading2 = (article) => {
    const title = article.children[0].children[0].displayTitle;
    nextHeading2Num++;
    nextHeading3Num = 0;
    // console.log(`${nextHeading1Num}.${nextHeading2Num}`);
    // console.log(title);
    createHeading(title, `${nextHeading1Num}.${nextHeading2Num}`, 2);
};

// generates heading 3 and contents underneath
const generateHeading3andContent = (teachingPoint) => {
    let content;
    const keyLearningPoints = teachingPoint.children;
    // console.log(keyLearningPoints);
    keyLearningPoints.forEach(child => child.children.forEach((object, index) => {
        if (index === 0) {
            nextHeading3Num++;
            createHeading(object.displayTitle, `${nextHeading1Num}.${nextHeading2Num}.${nextHeading3Num}`, 3);
        } else if (object._component === "ilt-text") {
            addText(object);
        } else if (object._component === "ilt-image") {
            addImage(object);

        } else if (object._component === "ilt-list") {
            addList(object);
        }
        // else if (object._component === "blank") {
        //     generateBlankPage();
        // }
    }));
};
const addText = (object) => {
    contents.push(new Paragraph({
        children: [
            new TextRun({
                // text: object.attributes.content,
                text: "What is Text area 101???"
            }),
        ],


    }));
};
const addList = (object) => {
    object.properties.listItems.forEach(listItem => { //it assumes all lists have bullet points - to FIX!!!
        contents.push(new Paragraph({
            children: [
                new TextRun({
                    text: listItem.textArea.replace(/<\/?[^>]+>/gi, '')
                })
            ],
            bullet: {
                level: 0
            }
        }));

    });
};
const addImage = (object) => {
    contents.push(new Paragraph({
        children: [
            new ImageRun({
                data: fs.readFileSync("./assets/helicopter-portrait.jpg"),
                // data: fs.readFileSync(object.properties.assetFile)
                transformation: {
                    width: 600,
                    height: 350,
                },

            }),
        ],
        spacing: {
            before: 200,
            after: 200,
        },
    }));
};

// const generateBlankPage = () => {
// contents.push(new Paragraph({
//     pageBreakBefore: true,
// }));
// };

const generateDocX = () => {
    const sectionChildren = [new TableOfContents("Summary", {
        hyperlink: true,
        headingStyleRange: "1-5",
    }), ...contents,];
    doc = new Document({
        sections: [{
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            children: [
                                new ImageRun({
                                    data: fs.readFileSync("./assets/lemonardo.jpg"),
                                    transformation: {
                                        width: 100,
                                        height: 55,
                                    },
                                }),
                                new TextRun({
                                    text: "LEMONARDO"
                                }),
                            ],
                            border: {
                                bottom: {
                                    color: "auto",
                                    space: 1,
                                    style: BorderStyle.SINGLE,
                                    size: 6,
                                },
                            },
                        })
                    ],
                }),

            },
            footers: {
                default: new Footer({
                    children: [new Paragraph({
                        children: [
                            new TextRun({
                                children: [PageNumber.CURRENT],
                            })
                        ],
                        alignment: AlignmentType.RIGHT,
                        border: {
                            top: {
                                color: "auto",
                                space: 1,
                                style: BorderStyle.SINGLE,
                                size: 6,
                            },
                        },

                    })],
                }),
            },
            features: {
                updateFields: true,
            },
            properties: {},
            children:
                sectionChildren,
        }],
    });

    Packer.toBuffer(doc).then((buffer) => {
        fs.writeFileSync("My Document.docx", buffer);
    });
};

const getHeadings = (course) => {
    course.children.forEach(page => {
        page.children.forEach((article, index) => {
            if (article._type === 'article') {
                console.log(index);
                if (index === 0) { // it means it's a first element before Lesson Introduction
                    //  What should I do with it? It has got tons of info to display
                    console.log("found ???");
                } else if (article._type === "article" && index === 1) { // it means it's a Lesson Introduction
                    //  What should I do with it? It has got tons of info to display    
                    console.log("found Lesson Introduction");
                }
                if (article.title === "Section Introduction") {
                    generateHeading1(article);
                }
                else if (article.title === "Title") {
                    generateHeading2(article);
                }
                else {
                    generateHeading3andContent(article);
                }
            }
            else {
                throw `${article.displayTitle} type is not an article`;
            }
        });
    });
    generateDocX();
};

// adds new heading to headings array
const createHeading = (text, number, headingLevel) => {
    let headinglvl;
    let indent;
    let pageBreak = false;
    switch (headingLevel) {
        case 1:
            headinglvl = HeadingLevel.HEADING_1;
            indent = 0;
            pageBreak = true;
            break;
        case 2:
            headinglvl = HeadingLevel.HEADING_2;
            indent = 200;
            pageBreak = previousHeading > 1;
            break;
        case 3:
            headinglvl = HeadingLevel.HEADING_3;
            indent = 400;
            pageBreak = previousHeading > 2;
            break;
    }
    previousHeading = headingLevel;
    contents.push(new Paragraph({
        children: [
            new TextRun({
                text: `${number} ${text}`,
            }),
        ],
        heading: headinglvl,
        indent: {
            left: indent,
        },
        pageBreakBefore: pageBreak,
        spacing: {
            after: 200,
        },
        style: paragraphStyles.heading1

    }));
};


createWordOutput();

module.exports = { createWordOutput, generateHeading1, generateHeading2, generateHeading3andContent, generateDocX, getHeadings, createHeading };