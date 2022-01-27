// Here goes all the styling and word output configuration

const docx = require("docx");
const { convertInchesToTwip, convertMillimetersToTwip, HeadingLevel } = require("docx");

const stylesConfig = module.exports = {
    // styles: {
    default: {
        document: {
            run: {
                font: "Arial",
            },
        },
    },
    // heading1: {
    //   run: {
    //     size: 28,
    //     bold: true,
    //     italics: true,
    //     color: "FF0000",
    //   },
    //   paragraph: {
    //     spacing: {
    //       after: 120,
    //     },
    //   },
    // },
    // heading2: {
    //   run: {
    //     size: 26,
    //     bold: true,
    //     underline: {
    //       type: UnderlineType.DOUBLE,
    //       color: "FF0000",
    //     },
    //   },
    //   paragraph: {
    //     spacing: {
    //       before: 240,
    //       after: 120,
    //     },
    //   },
    // },
    paragraphStyles: [
        {
            id: "WorkbitHeading1",
            name: "WorkbitHeading1",
            basedOn: "Heading1",
            next: "Heading1",
            quickFormat: true,
            paragraph: {
                heading: HeadingLevel.HEADING_1,
            },
            run: {
                font: "Arial",
                size: 28,
                bold: true,
                color: "#000000",
            },
        },
        {
            id: "WorkbitHeading2",
            name: "WorkbitHeading2",
            basedOn: "Heading2",
            next: "Heading2",
            quickFormat: true,
            paragraph: {
                heading: HeadingLevel.HEADING_2,
            },
            run: {
                font: "Arial",
                size: 24,
                bold: true,
                color: "#000000",
            },
        },
        {
            id: "WorkbitHeading3",
            name: "WorkbitHeading3",
            basedOn: "Heading3",
            next: "Heading3",
            quickFormat: true,
            paragraph: {
                heading: HeadingLevel.HEADING_3,
            },
            run: {
                font: "Arial",
                size: 20,
                bold: true,
                color: "#000000",
            },
        },
        {
            id: "normalPara",
            name: "Normal Para",
            basedOn: "Normal",
            next: "Normal",
            quickFormat: true,
            run: {
                font: "Arial",
                size: 20,
                // bold: true,de
                // color: "#FF0000"
            },
            paragraph: {
                spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
                // rightTabStop: TabStopPosition.MAX,
                // leftTabStop: 453.543307087,
            },
        },
    ],
    // },
};

module.exports = stylesConfig;