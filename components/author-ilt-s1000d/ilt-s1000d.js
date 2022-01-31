const docx = require("docx");
const { Paragraph, TextRun, ExternalHyperlink, AlignmentType } = require("docx");

const adds1000d = module.exports = (object) => {
    const newText = [];

    let audio;
    const objProps = object.properties;
    if (objProps.preTextArea.replace(/<\/?[^>]+>/gi, '') || objProps.textAreaField.replace(/<\/?[^>]+>/gi, '') || objProps.postTextArea.replace(/<\/?[^>]+>/gi, '') || objProps.audio.replace(/<\/?[^>]+>/gi, '')) {
        if (object.properties.audio) {
            audio = new ExternalHyperlink({
                children: [
                    new TextRun({
                        text: "S1000D audio file",
                        style: "Hyperlink",
                        break: 1,
                    }),
                ],
                link: "./assets/audio.mp3",
                // link: object.properties.audio,
            });
        }

        newText.push(new Paragraph({
            children: [
                new TextRun({
                    text: object.properties.displayType.replace(/<\/?[^>]+>/gi, '').toUpperCase()
                })
            ],
            alignment: AlignmentType.CENTER,
            style: object.title,
            keepNext: true,
        }));

        newText.push(new Paragraph({
            children: [
                new TextRun({
                    text: object.properties.preTextArea.replace(/<\/?[^>]+>/gi, ''),
                }),
                new TextRun({
                    text: object.properties.textAreaField.replace(/<\/?[^>]+>/gi, ''),
                    break: 1,
                }),
                audio,
                new TextRun({
                    text: object.properties.postTextArea.replace(/<\/?[^>]+>/gi, ''),
                    break: 1,
                }),
            ],
            style: object.title,
            widowControl: true,
            keepLines: true
        }));
        return newText;
    }
    else {
        console.log("No properties or assets for: " + object._id);
        return [];
    };
};

module.exports = adds1000d;