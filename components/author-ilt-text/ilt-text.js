const docx = require("docx");
const { Paragraph, TextRun } = require("docx");

const addText = (module.exports = (object) => {
  if (object.displayTitle) {
    return new Paragraph({
      children: [
        new TextRun({
          text: object.displayTitle,
        }),
      ],
      style: "textPara",
      keepLines: true,
    });
  } else {
    console.log("No text for: " + object._id);
  }
});

module.exports = addText;
