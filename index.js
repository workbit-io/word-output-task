const docx = require("docx");
const fs = require("fs");
const data = require("./wlodek_json_word.json");
const { Document, Packer, Paragraph, TextRun, HeadingLevel } = require("docx");

let doc;
const headings = [];

// usefull to keep the track of heading numbers
let nextHeading1Num = 0;
let nextHeading2Num = 0;
let nextHeading3Num = 0;

const createWordOutput = () => {
    const course = data.content[0];
    generateHeadings(course);
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
    console.log(nextHeading1Num);
    console.log(sessionIntroductionTitle[0].DisplayTitle);
    createHeading(sessionIntroductionTitle[0].DisplayTitle, nextHeading1Num, 1);
};

const generateHeading2 = (article) => {
    const title = article.children[0].children[0].DisplayTitle;
    nextHeading2Num++;
    nextHeading3Num = 0;
    console.log(`${nextHeading1Num}.${nextHeading2Num}`);
    console.log(title);
    createHeading(title, `${nextHeading1Num}.${nextHeading2Num}`, 2);
};

const generateHeading3 = (teachingPoint) => {
    const keyLearningPoints = teachingPoint.children.map(block => block);
    const headings = keyLearningPoints.map(KLP => KLP.children[0].DisplayTitle);
    headings.forEach(heading => {
        nextHeading3Num++;
        console.log(`${nextHeading1Num}.${nextHeading2Num}.${nextHeading3Num}`);
        console.log(heading);
        createHeading(heading, `${nextHeading1Num}.${nextHeading2Num}.${nextHeading3Num}`, 3);
    });
};

const generateDocX = () => {
    // console.log(headings);

    doc = new Document({
        sections: [{
            properties: {},
            children: headings,
        }],
    });

    Packer.toBuffer(doc).then((buffer) => {
        fs.writeFileSync("My Document.docx", buffer);
    });
};

const generateHeadings = (course) => {
    course.children.forEach(page => {
        page.children.forEach(article => {
            if (article.title === "Section Introduction") {
                // console.log("Section Introduction: ");
                generateHeading1(article);
            }
            else if (article.title === "Title") {
                // console.log("Title: ");
                generateHeading2(article);
            }
            else {
                // console.log("Teaching point: ");
                generateHeading3(article);
            }
        });
    });
    generateDocX();
};

// adds new heading to headings array
const createHeading = (text, number, headingLevel) => {
    let headinglvl;
    switch (headingLevel) {
        case 1:
            headinglvl = HeadingLevel.HEADING_1;
            break;
        case 2:
            headinglvl = HeadingLevel.HEADING_2;
            break;
        case 3:
            headinglvl = HeadingLevel.HEADING_3;
            break;
    }
    headings.push(new Paragraph({
        children: [
            new TextRun({
                text: `${number} ${text}`,
            }),
        ],
        heading: headinglvl
    }));
};


// generateDocX();

createWordOutput();