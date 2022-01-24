const docx = require("docx");
const fs = require("fs");
const data = require("./wlodek_json_word.json");
const { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel, TableOfContents } = require("docx");

let doc;
const contents = [];
let previousHeading = 1;

// keep the track of heading numbering
let nextHeading1Num = 0;
let nextHeading2Num = 0;
let nextHeading3Num = 0;

// extracts course from the json file
const createWordOutput = () => {
    const course = data.content[0];
    getHeadings(course);
};

const generateHeading1 = (article) => {
    const sessionIntroductionTitle = article.children[0].children.filter(object => {
        return object.title === "Image Element Label";
    });
    // setting next HEADING numbering to +1
    // resetting HEADING 2 and HEADING 3 as with every article/Section Introduction element numbering starts from beginning
    nextHeading1Num++;
    nextHeading2Num = 0;
    nextHeading3Num = 0;
    // console.log(nextHeading1Num);
    // console.log(sessionIntroductionTitle[0].DisplayTitle);
    createHeading(sessionIntroductionTitle[0].DisplayTitle, nextHeading1Num, 1);
};

const generateHeading2 = (article) => {
    const title = article.children[0].children[0].DisplayTitle;
    nextHeading2Num++;
    nextHeading3Num = 0;
    // console.log(`${nextHeading1Num}.${nextHeading2Num}`);
    // console.log(title);
    createHeading(title, `${nextHeading1Num}.${nextHeading2Num}`, 2);
};

// add filtering to find KLP's title
// const generateHeading3andContent = (teachingPoint) => {
//     const keyLearningPoints = teachingPoint.children;
//     const headings = keyLearningPoints.map(KLP => KLP.children[0].DisplayTitle);
//     headings.forEach(heading => {
//         nextHeading3Num++;
//         // console.log(`${nextHeading1Num}.${nextHeading2Num}.${nextHeading3Num}`);
//         // console.log(heading);
//         createHeading(heading, `${nextHeading1Num}.${nextHeading2Num}.${nextHeading3Num}`, 3);
//     });
//     // generateContent(KeyLearningPoints);
// };

// generates heading 3 and contents underneath
const generateHeading3andContent = (teachingPoint) => {
    let content;
    const keyLearningPoints = teachingPoint.children;
    // console.log(keyLearningPoints);
    keyLearningPoints.forEach(child => child.children.forEach((object, index) => {
        if (index === 0) {
            nextHeading3Num++;
            createHeading(object.DisplayTitle, `${nextHeading1Num}.${nextHeading2Num}.${nextHeading3Num}`, 3);
        } else if (object.componentType === "ilt-text") {
            contents.push(new Paragraph({
                children: [
                    new TextRun({
                        text: object.attributes.content,
                    }),
                ],


            }));
        } else if (object.componentType === "ilt-image") {
            contents.push(new Paragraph({
                children: [
                    new ImageRun({
                        data: fs.readFileSync("./assets/helicopter.jpg"),
                        transformation: {
                            width: 600,
                            height: 750,
                        },
                    }),
                ],
            }));
        }
    }));
};

const generateDocX = () => {
    const sectionChildren = [new TableOfContents("Summary", {
        hyperlink: true,
        headingStyleRange: "1-5",
    }), ...contents,];
    doc = new Document({
        sections: [{
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
        page.children.forEach(article => {
            if (article.type === 'article') {
                if (article.title === "Section Introduction") {
                    generateHeading1(article);
                }
                else if (article.title === "Title") {
                    generateHeading2(article);
                }
                else {
                    generateHeading3andContent(article);
                }
            } else {
                throw `${article.DisplayTitle} type is not an article`;
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

    }));
};


createWordOutput();

module.exports = { createWordOutput, generateHeading1, generateHeading2, generateHeading3andContent, generateDocX, getHeadings, createHeading };