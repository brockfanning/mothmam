const { Document, AlignmentType, Paragraph, Packer, TextRun } = require('docx')
const fs = require('fs')

const doc = new Document();

doc.addSection({
    children: [
        new Paragraph({
            bullet: {
                level: 0,
            },
            children: [
                new TextRun({
                    italics: true,
                    bold: true,
                    text: 'Foo'
                }),
                new TextRun('Bar')
            ]
        }),
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("example.docx", buffer);
});