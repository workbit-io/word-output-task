// Here goes all the styling and word output configuration

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
            run: {
                font: "Arial",
                size: 28,
                bold: true,
                color: "#FF0000",
            },
        },
        {
            id: "WorkbitHeading2",
            name: "WorkbitHeading2",
            basedOn: "Heading2",
            next: "Heading2",
            quickFormat: true,
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
            run: {
                font: "Arial",
                size: 20,
                bold: true,
                color: "#000000",
            },
        },
    ],
    // },
};

module.exports = stylesConfig;