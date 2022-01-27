const docx = require("docx");
const fs = require("fs");
const { Paragraph, ImageRun } = require("docx");

const addImage = module.exports = (object) => {
    return (new Paragraph({
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
        style: "imagePara"
    }));

};

module.exports = addImage;