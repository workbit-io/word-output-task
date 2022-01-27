const docx = require("docx");
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel, TableOfContents, BorderStyle, FrameAnchorType, HorizontalPositionAlign, VerticalPositionAlign, LevelFormat, UnderlineType, AlignmentType, TabStopPosition, convertInchesToTwip, convertMillimetersToTwip, Footer, ExternalHyperlink } = require("docx");
const stylesConfig = require("./config");

// console.log(stylesConfig.paragraphStyles.find(style => style.id === "WorkbitHeading1").run);
// const headingStyle = stylesConfig.paragraphStyles.find(style => style.id === "WorkbitHeading1").run;

const doc = new Document({
    // styles: stylesConfig,
    sections: [{

        children: [
            new Paragraph({
                children: [
                    new ExternalHyperlink({
                        children: [
                            new TextRun({
                                text: "This is an external link!",
                                style: "Hyperlink",
                            }),
                        ],
                        // link: "./video.mp4",
                        link: "./assets/video.mp4",
                    }),
                ],
            })]
    }
    ]
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My test Document.docx", buffer);
});

var dir = './needed-assets';

        // crestes dir
// if (!fs.existsSync(dir)) {
//     fs.mkdirSync(dir);
// }

        // copies file
// fs.copyFile('source.txt', 'destination.txt', (err) => {
//     if (err) throw err;
//     console.log('source.txt was copied to destination.txt');
//   });

        // copies file to new destination
//   fs.copySync(path.resolve(assets,'./mainisp.jpg'), './test/mainisp.jpg');