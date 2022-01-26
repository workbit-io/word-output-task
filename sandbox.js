const docx = require("docx");
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel, TableOfContents, BorderStyle, FrameAnchorType, HorizontalPositionAlign, VerticalPositionAlign, LevelFormat, UnderlineType, AlignmentType, TabStopPosition, convertInchesToTwip, convertMillimetersToTwip, Footer } = require("docx");
const stylesConfig = require("./config");

// console.log(stylesConfig.paragraphStyles.find(style => style.id === "WorkbitHeading1").run);
// const headingStyle = stylesConfig.paragraphStyles.find(style => style.id === "WorkbitHeading1").run;

const doc = new Document({
    styles: stylesConfig,
    // styles: {
    //     default: {
    //         document: {
    //             run: {
    //                 font: "Arial",
    //             },
    //         },
    //     },
    //     // heading1: {
    //     //   run: {
    //     //     size: 28,
    //     //     bold: true,
    //     //     italics: true,
    //     //     color: "FF0000",
    //     //   },
    //     //   paragraph: {
    //     //     spacing: {
    //     //       after: 120,
    //     //     },
    //     //   },
    //     // },
    //     // heading2: {
    //     //   run: {
    //     //     size: 26,
    //     //     bold: true,
    //     //     underline: {
    //     //       type: UnderlineType.DOUBLE,
    //     //       color: "FF0000",
    //     //     },
    //     //   },
    //     //   paragraph: {
    //     //     spacing: {
    //     //       before: 240,
    //     //       after: 120,
    //     //     },
    //     //   },
    //     // },
    //     paragraphStyles: [
    //         {
    //             id: "WorkbitHeading1",
    //             name: "WorkbitHeading1",
    //             basedOn: "Heading1",
    //             next: "Heading1",
    //             quickFormat: true,
    //             run: {
    //                 font: "Calibri",
    //                 size: 28,
    //                 bold: true,
    //                 color: "#FF0000",
    //             },
    //         },
    //         {
    //             id: "WorkbitHeading2",
    //             name: "WorkbitHeading2",
    //             basedOn: "Heading2",
    //             next: "Heading2",
    //             quickFormat: true,
    //             run: {
    //                 font: "Arial",
    //                 size: 24,
    //                 bold: true,
    //                 color: "#000000",
    //             },
    //         },
    //         {
    //             id: "WorkbitHeading3",
    //             name: "WorkbitHeading3",
    //             basedOn: "Heading3",
    //             next: "Heading3",
    //             quickFormat: true,
    //             run: {
    //                 font: "Arial",
    //                 size: 20,
    //                 bold: true,
    //                 color: "#000000",
    //             },
    //         },
    //     ],
    // },

    sections: [
        {
            children: [
                new Paragraph("No border!"),
                new Paragraph({
                    style: "WorkbitHeading1",
                    text: "I have borders on my top and bottom sides!",
                    border: {
                        top: {
                            color: "auto",
                            space: 1,
                            style: BorderStyle.SINGLE,
                            size: 6,
                        },
                        bottom: {
                            color: "auto",
                            space: 1,
                            style: BorderStyle.SINGLE,
                            size: 6,
                        },
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "This will ",
                        }),
                        new TextRun({
                            text: "have a border.",
                            border: {
                                color: "auto",
                                space: 1,
                                style: BorderStyle.SINGLE,
                                size: 6,
                            },
                        }),
                        new TextRun({
                            text: " This will not.",
                        }),
                    ],
                }),
                new Paragraph({
                    frame: {
                        position: {
                            x: 1000,
                            y: 3000,
                        },
                        width: 4000,
                        height: 1000,
                        anchor: {
                            horizontal: FrameAnchorType.MARGIN,
                            vertical: FrameAnchorType.MARGIN,
                        },
                        alignment: {
                            x: HorizontalPositionAlign.CENTER,
                            y: VerticalPositionAlign.TOP,
                        },
                    },
                    border: {
                        top: {
                            color: "auto",
                            space: 1,
                            value: "single",
                            size: 6,
                        },
                        bottom: {
                            color: "auto",
                            space: 1,
                            value: "single",
                            size: 6,
                        },
                        left: {
                            color: "auto",
                            space: 1,
                            value: "single",
                            size: 6,
                        },
                        right: {
                            color: "auto",
                            space: 1,
                            value: "single",
                            size: 6,
                        },
                    },
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({
                            text: "Foo Bar",
                            bold: true,
                        }),
                        new TextRun({
                            text: "\tGithub is the best",
                            bold: true,
                        }),
                    ],
                })
            ],

        },
    ],
});

// const doc = new Document({
//     numbering: {
//         config: [{
//             reference: 'ref1',
//             levels: [
//                 {
//                     level: 0,
//                     format: LevelFormat.DECIMAL,
//                     text: '%1)',
//                     start: 50,
//                 }
//             ],
//         }]
//     },
//     styles: {
//         default: {
//             heading1: {
//                 run: {
//                     font: "Calibri",
//                     size: 52,
//                     bold: true,
//                     color: "000000",
//                     underline: {
//                         type: UnderlineType.SINGLE,
//                         color: "000000",
//                     },
//                 },
//                 paragraph: {
//                     alignment: AlignmentType.CENTER,
//                     spacing: { line: 340 },
//                 },
//             },
//             heading2: {
//                 run: {
//                     font: "Calibri",
//                     size: 26,
//                     bold: true,
//                 },
//                 paragraph: {
//                     spacing: { line: 340 },
//                 },
//             },
//             heading3: {
//                 run: {
//                     font: "Calibri",
//                     size: 26,
//                     bold: true,
//                 },
//                 paragraph: {
//                     spacing: { line: 276 },
//                 },
//             },
//             heading4: {
//                 run: {
//                     font: "Calibri",
//                     size: 26,
//                     bold: true,
//                 },
//                 paragraph: {
//                     alignment: AlignmentType.JUSTIFIED,
//                 },
//             },
//         },
//         paragraphStyles: [
//             {
//                 id: "normalPara",
//                 name: "Normal Para",
//                 basedOn: "Normal",
//                 next: "Normal",
//                 quickFormat: true,
//                 run: {
//                     font: "Calibri",
//                     size: 26,
//                     bold: true,
//                 },
//                 paragraph: {
//                     spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
//                     rightTabStop: TabStopPosition.MAX,
//                     leftTabStop: 453.543307087,
//                 },
//             },
//             {
//                 id: "normalPara2",
//                 name: "Normal Para2",
//                 basedOn: "Normal",
//                 next: "Normal",
//                 quickFormat: true,
//                 run: {
//                     font: "Calibri",
//                     size: 26,
//                 },
//                 paragraph: {
//                     alignment: AlignmentType.JUSTIFIED,
//                     spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
//                 },
//             },
//             {
//                 id: "aside",
//                 name: "Aside",
//                 basedOn: "Normal",
//                 next: "Normal",
//                 run: {
//                     color: "999999",
//                     italics: true,
//                 },
//                 paragraph: {
//                     spacing: { line: 276 },
//                     indent: { left: convertInchesToTwip(0.5) },
//                 },
//             },
//             {
//                 id: "wellSpaced",
//                 name: "Well Spaced",
//                 basedOn: "Normal",
//                 paragraph: {
//                     spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
//                 },
//             },
//             {
//                 id: "numberedPara",
//                 name: "Numbered Para",
//                 basedOn: "Normal",
//                 next: "Normal",
//                 quickFormat: true,
//                 run: {
//                     font: "Calibri",
//                     size: 26,
//                     bold: true,
//                 },
//                 paragraph: {
//                     spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
//                     rightTabStop: TabStopPosition.MAX,
//                     leftTabStop: 453.543307087,
//                     numbering: {
//                         reference: 'ref1',
//                         instance: 0,
//                         level: 0,
//                     }
//                 },
//             },
//         ],
//     },
//     sections: [
//         {
//             properties: {
//                 page: {
//                     margin: {
//                         top: 700,
//                         right: 700,
//                         bottom: 700,
//                         left: 700,
//                     },
//                 },
//             },
//             footers: {
//                 default: new Footer({
//                     children: [
//                         new Paragraph({
//                             text: "1",
//                             style: "normalPara",
//                             alignment: AlignmentType.RIGHT,
//                         }),
//                     ],
//                 }),
//             },
//             children: [
//                 // new Paragraph({
//                 //     children: [
//                 //         new ImageRun({
//                 //             data: fs.readFileSync(".assets/helicopter-portrait.jpg"),
//                 //             transformation: {
//                 //                 width: 100,
//                 //                 height: 100,
//                 //             },
//                 //         }),
//                 //     ],
//                 // }),
//                 new Paragraph({
//                     text: "HEADING",
//                     heading: HeadingLevel.HEADING_1,
//                     alignment: AlignmentType.CENTER,
//                 }),
//                 new Paragraph({
//                     text: "Ref. :",
//                     style: "normalPara",
//                 }),
//                 new Paragraph({
//                     text: "Date :",
//                     style: "normalPara",
//                 }),
//                 new Paragraph({
//                     text: "To,",
//                     style: "normalPara",
//                 }),
//                 new Paragraph({
//                     text: "The Superindenting Engineer,(O &M)",
//                     style: "normalPara",
//                 }),
//                 new Paragraph({
//                     text: "Sub : ",
//                     style: "normalPara",
//                 }),
//                 new Paragraph({
//                     text: "Ref. : ",
//                     style: "normalPara",
//                 }),
//                 new Paragraph({
//                     text: "Sir,",
//                     style: "normalPara",
//                 }),
//                 new Paragraph({
//                     text: "BRIEF DESCRIPTION",
//                     style: "normalPara",
//                 }),
//                 // table,
//                 // new Paragraph({
//                 //     children: [
//                 //         new ImageRun({
//                 //             data: fs.readFileSync("./assets/helicopter-portrait.jpg"),
//                 //             transformation: {
//                 //                 width: 100,
//                 //                 height: 100,
//                 //             },
//                 //         }),
//                 //     ],
//                 // }),
//                 new Paragraph({
//                     text: "Test",
//                     style: "normalPara2",
//                 }),
//                 // new Paragraph({
//                 //     children: [
//                 //         new ImageRun({
//                 //             data: fs.readFileSync("./assets/helicopter-portrait.jpg"),
//                 //             transformation: {
//                 //                 width: 100,
//                 //                 height: 100,
//                 //             },
//                 //         }),
//                 //     ],
//                 // }),
//                 new Paragraph({
//                     text: "Test 2",
//                     style: "normalPara2",
//                 }),
//                 new Paragraph({
//                     text: "Numbered paragraph that has numbering attached to custom styles",
//                     style: "numberedPara",
//                 }),
//                 new Paragraph({
//                     text: "Numbered para would show up in the styles pane at Word",
//                     style: "numberedPara",
//                 }),
//             ],
//         },
//     ],
// });

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My test Document.docx", buffer);
});
