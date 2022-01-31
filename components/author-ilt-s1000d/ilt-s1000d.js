const docx = require("docx");
const { Paragraph, TextRun, ExternalHyperlink, AlignmentType } = require("docx");

const adds1000d = module.exports = (object) => {
    const newText = [];
    const title = object.displayTitle;

    let audio;
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
    }));

    newText.push(new Paragraph({
        children: [
            // new TextRun({
            //     text: title,
            // }),
            // new TextRun({
            //     text: object.properties.displayType.replace(/<\/?[^>]+>/gi, '').toUpperCase(),
            //     alignment: AlignmentType.CENTER
            //     // break: 1,
            // }),
            new TextRun({
                text: object.properties.preTextArea.replace(/<\/?[^>]+>/gi, ''),
                // break: 1,
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
    }));


    return newText;
};

module.exports = adds1000d;