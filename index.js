const docx = require("docx");
const fs = require("fs");
const data = require("./telebrief.json");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ImageRun,
  TableOfContents,
  Header,
  Footer,
  PageNumber,
  AlignmentType,
  BorderStyle,
} = require("docx");
const stylesConfig = require("./config");
const addIltList = require("./components/author-ilt-list/ilt-list");
const addIltText = require("./components/author-ilt-text/ilt-text");
const addIltImage = require("./components/author-ilt-image/ilt-image");
const addIltAV = require("./components/author-ilt-av/ilt-av");
const addIltAnimate = require("./components/author-ilt-animate/ilt-animate");
const addIltMcq = require("./components/author-ilt-mcq/ilt-mcq");
const addIltDragImage = require("./components/author-ilt-drag-image/ilt-drag-image");
const addIlts1000d = require("./components/author-ilt-s1000d/ilt-s1000d");

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
  const sessionIntroductionTitle = article.children[0].children[0];
  // setting next HEADING numbering to +1
  // resetting HEADING 2 and HEADING 3 as with every article/Section Introduction element numbering starts from beginning
  nextHeading1Num++;
  nextHeading2Num = 0;
  nextHeading3Num = 0;
  createHeading(sessionIntroductionTitle.displayTitle, nextHeading1Num, 1);
};

const generateHeading2 = (article) => {
  const title = article.children[0].children[0].displayTitle;
  nextHeading2Num++;
  nextHeading3Num = 0;
  createHeading(title, `${nextHeading1Num}.${nextHeading2Num}`, 2);
};

// generates heading 3 and contents underneath
const generateHeading3andContent = (teachingPoint) => {
  const keyLearningPoints = teachingPoint.children;
  keyLearningPoints.forEach((child) =>
    child.children.forEach((object, index) => {
      if (index === 0) {
        nextHeading3Num++;
        createHeading(
          object.displayTitle,
          `${nextHeading1Num}.${nextHeading2Num}.${nextHeading3Num}`,
          3
        );
        // checks if heading 3 contains image
        if (object.properties.assetFile) {
          // addImage(object);
          // placeholder to replace for production with addImage() function
          contents.push(
            new Paragraph({
              children: [
                new ImageRun({
                  data: fs.readFileSync("./assets/heading3.jpg"),
                  // data: fs.readFileSync(object.properties.assetFile)
                  transformation: {
                    width: 600,
                    height: 250,
                  },
                }),
              ],
              style: "imagePara",
            })
          );
        }
      } else if (object._component === "ilt-text") {
        addText(object);
      } else if (object._component === "ilt-image") {
        addImage(object);
      } else if (object._component === "ilt-list") {
        addList(object);
      } else if (object._component === "ilt-av") {
        addAV(object);
      } else if (object._component === "ilt-animate") {
        addAnimate(object);
      } else if (object._component === "ilt-mcq") {
        addMcq(object);
      } else if (object._component === "ilt-drag-image") {
        addDragImage(object);
      } else if (object._component === "ilt-s1000d") {
        adds1000d(object);
      } else if (object._component === "blank") {
        generateBlankPage();
      }
    })
  );
};
const addText = (object) => {
  contents.push(addIltText(object));
};

const addList = (object) => {
  addIltList(object) &&
    addIltList(object).forEach((element) => contents.push(element));
};

const addImage = (object) => {
  contents.push(addIltImage(object));
};

const addAV = (object) => {
  contents.push(addIltAV(object));
};

const addAnimate = (object) => {
  contents.push(addIltAnimate(object));
};

const addMcq = (object) => {
  addIltMcq(object) &&
    addIltMcq(object).forEach((element) => contents.push(element));
};

const addDragImage = (object) => {
  addIltDragImage(object) &&
    addIltDragImage(object).forEach((element) => contents.push(element));
};

const adds1000d = (object) => {
  addIlts1000d(object) &&
    addIlts1000d(object).forEach((element) => contents.push(element));
};
const generateBlankPage = () => {
  contents.push(
    new Paragraph({
      pageBreakBefore: true,
    })
  );
};

const generateDocX = () => {
  // adds a table of contents to the doc
  const sectionChildren = [
    new TableOfContents("Table of contents", {
      hyperlink: true,
      headingStyleRange: "1-5",
    }),
    ...contents,
  ];

  doc = new Document({
    styles: stylesConfig, // gets styling from config.js
    sections: [
      {
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
                    text: "LEMONARDO",
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
                spacing: {
                  after: 200,
                },
              }),
            ],
          }),
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    children: [PageNumber.CURRENT],
                  }),
                ],
                alignment: AlignmentType.RIGHT,
                border: {
                  top: {
                    color: "auto",
                    space: 50,
                    style: BorderStyle.SINGLE,
                    size: 6,
                  },
                },
                spacing: {
                  after: 200,
                },
              }),
            ],
          }),
        },
        // enables table of contents functionality
        features: {
          updateFields: true,
        },
        properties: {},
        children: sectionChildren, // all children and table of contents
      },
    ],
  });

  Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
  });
};

const getHeadings = (course) => {
  course.children.forEach((page) => {
    page.children.forEach((article, index) => {
      if (article._type === "article") {
        if (index === 0) {
          // it means it's a first element before Lesson Introduction
          //  What should I do with it? It has got tons of info to display
          // console.log("found and omitted ???");
          return;
        } else if (index === 1) {
          // it means it's a Lesson Introduction
          //  What should I do with it? It has got tons of info to display
          // console.log("found and omitted Lesson Introduction");
          return;
        }
        if (article.title === "Section Introduction") {
          generateHeading1(article);
        } else if (article.title === "Title") {
          generateHeading2(article);
        } else {
          generateHeading3andContent(article);
        }
      } else {
        throw `${article.displayTitle} type is not an article`;
      }
    });
  });
  generateDocX();
};

// adds new heading to headings array
const createHeading = (text, number, headingLevel) => {
  let pageBreak = false;
  switch (headingLevel) {
    case 1:
      pageBreak = true;
      break;
    case 2:
      pageBreak = previousHeading > 1;
      break;
    case 3:
      pageBreak = previousHeading > 2;
      break;
  }
  previousHeading = headingLevel;
  contents.push(
    new Paragraph({
      style: `WorkbitHeading${headingLevel}`,
      children: [
        new TextRun({
          text: `${number} ${text}`,
        }),
      ],
      pageBreakBefore: pageBreak,
      spacing: {
        after: 200,
      },
    })
  );
};

createWordOutput();

// for testing purposes
module.exports = {
  createWordOutput,
  generateHeading1,
  generateHeading2,
  generateHeading3andContent,
  generateDocX,
  getHeadings,
  createHeading,
};
