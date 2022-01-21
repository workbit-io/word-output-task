// import * as fs from "fs";
const fs = require("fs");

const { Document, Packer, Paragraph, TextRun } = require("docx");
const { builtinModules } = require("module");

// Documents contain sections, you can have multiple sections per document, go here to learn more about sections
// This simple example will only contain one section
const doc = new Document({
    sections: [{
        properties: {},
        children: [
            new Paragraph({
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
            }),
        ],
    }],
});

// Used to export the file into a .docx file
Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});

module.exports = { doc };

// Done! A file called 'My Document.docx' will be in your file system.