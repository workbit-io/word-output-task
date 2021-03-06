// Here goes all the styling and word output configuration

const docx = require("docx");
const { convertInchesToTwip, convertMillimetersToTwip, HeadingLevel, BorderStyle } = require("docx");

const stylesConfig = module.exports = {
    default: {
        document: {
            run: {
                font: "Arial",
            },
        },
    },
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
            },
            paragraph: {
                spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
            },
        },
        {
            id: "bulletPara",
            name: "Bullet Para",
            basedOn: "Bullet",
            next: "Bullet",
            quickFormat: true,
            run: {
                font: "Arial",
                size: 20,
            },
            paragraph: {
                spacing: {
                    line: 276,
                    before: 20 * 72 * 0.1,
                    after: 20 * 72 * 0.05
                },
                bullet: {
                    level: 0
                },
            },
        },
        {
            id: "imagePara",
            name: "Image Para",
            basedOn: "Image",
            next: "Image",
            quickFormat: true,
            paragraph: {
                spacing: {
                    before: 200,
                    after: 200
                }
            }
        },
        {
            id: "greyedOutPara",
            name: "Greyed Out Para",
            basedOn: "Greyed Out",
            next: "Greyed Out",
            quickFormat: true,
            run: {
                font: "Arial",
                size: 20,
                color: "#808080",
            },
            paragraph: {
                spacing: {
                    line: 276,
                    before: 20 * 72 * 0.1,
                    after: 20 * 72 * 0.05
                },
            }
        },
        {
            id: "textPara",
            name: "Text Para",
            basedOn: "Text",
            next: "Text",
            quickFormat: true,
            run: {
                font: "Arial",
                size: 20,
                color: "auto"
            },
            paragraph: {
                spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
                border: {
                    top: {
                        color: "auto",
                        space: 1,
                        style: BorderStyle.DASHED,
                        size: 15,
                    },
                    bottom: {
                        color: "auto",
                        space: 1,
                        style: BorderStyle.DASHED,
                        size: 15,
                    },
                    left: {
                        color: "auto",
                        space: 10,
                        style: BorderStyle.DASHED,
                        size: 15,
                    },
                    right: {
                        color: "auto",
                        space: 10,
                        style: BorderStyle.DASHED,
                        size: 15,
                    },
                },
            },
        },
        {
            id: "S1000D Caution",
            name: "S1000D Caution",
            basedOn: "S1000d",
            next: "S1000d",
            quickFormat: true,
            run: {
                font: "Arial",
                size: 20,
                color: "auto",
                bold: true,
            },
            paragraph: {
                spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
                border: {
                    top: {
                        color: "#FFFF00",
                        space: 1,
                        style: BorderStyle.DASH_DOT_STROKED,
                        size: 50,
                    },
                    bottom: {
                        color: "#FFFF00",
                        space: 1,
                        style: BorderStyle.DASH_DOT_STROKED,
                        size: 50,
                    },
                    left: {
                        color: "#FFFF00",
                        space: 1,
                        style: BorderStyle.DASH_DOT_STROKED,
                        size: 50,
                    },
                    right: {
                        color: "#FFFF00",
                        space: 1,
                        style: BorderStyle.DASH_DOT_STROKED,
                        size: 50,
                    },
                },
            },
        },
        {
            id: "S1000D Warning",
            name: "S1000D Warning",
            basedOn: "S1000d",
            next: "S1000d",
            quickFormat: true,
            run: {
                font: "Arial",
                size: 20,
                color: "auto",
                bold: true,
            },
            paragraph: {
                spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
                border: {
                    top: {
                        color: "#FF0000",
                        space: 1,
                        style: BorderStyle.DASH_DOT_STROKED,
                        size: 50,
                    },
                    bottom: {
                        color: "#FF0000",
                        space: 1,
                        style: BorderStyle.DASH_DOT_STROKED,
                        size: 50,
                    },
                    left: {
                        color: "#FF0000",
                        space: 1,
                        style: BorderStyle.DASH_DOT_STROKED,
                        size: 50,
                    },
                    right: {
                        color: "#FF0000",
                        space: 1,
                        style: BorderStyle.DASH_DOT_STROKED,
                        size: 50,
                    },
                },
            },
        },
        {
            id: "S1000D None",
            name: "S1000D None",
            basedOn: "S1000d",
            next: "S1000d",
            quickFormat: true,
            run: {
                font: "Arial",
                size: 20,
                color: "auto",
                bold: true,
            },
            paragraph: {
                spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
                border: {
                    top: {
                        color: "auto",
                        space: 1,
                        style: BorderStyle.DASH_DOT_STROKED,
                        size: 50,
                    },
                    bottom: {
                        color: "auto",
                        space: 1,
                        style: BorderStyle.DASH_DOT_STROKED,
                        size: 50,
                    },
                    left: {
                        color: "auto",
                        space: 1,
                        style: BorderStyle.DASH_DOT_STROKED,
                        size: 50,
                    },
                    right: {
                        color: "auto",
                        space: 1,
                        style: BorderStyle.DASH_DOT_STROKED,
                        size: 50,
                    },
                },
            },
        },
    ],
};

module.exports = stylesConfig;