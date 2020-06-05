const cheerio = require('cheerio')
const { Paragraph, TextRun, Document, Packer, HeadingLevel} = require('docx')
const fs = require('fs')
const util = require('util')

let runType = 'debug'
runType = 'execute'
runType = 'contrived'

function debug(foo) {
    console.log(util.inspect(foo, false, null, true))
}

const testHtml = `
<h1>Hello <em>world</em></h1>
<ul>
    <li>One</li>
    <li>Two</li>
</ul>
`

const contrivedSections = [
    new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [
            new TextRun({
                text: 'Hello '
            }),
            new TextRun({
                italics: true,
                text: 'world'
            })
        ]
    }),
    new Paragraph({
        text: 'Foo',
        bullet: { level: 0 }
    }),
    new Paragraph({
        text: 'Bar',
        bullet: { level: 0 }
    })
]

const $ = cheerio.load(testHtml)
const doc = new Document()
const sections = []

$('body').contents('*').each((idx, element) => {
    const converted = convertElement(element)
    if (converted) {
        sections.push(converted)
    }
})

if (runType === 'debug') {
    debug(sections)
}
if (runType === 'execute') {
    doc.addSection({
        properties: {},
        children: sections
    })
    Packer.toBuffer(doc).then((buffer) => {
        fs.writeFileSync('example.docx', buffer);
    });
}
if (runType === 'contrived') {
    doc.addSection({
        properties: {},
        children: contrivedSections,
    })
    Packer.toBuffer(doc).then((buffer) => {
        fs.writeFileSync('example.docx', buffer);
    });
}


function convertElement(element, objectOnly=false) {

    let skip = false
    let type = null
    const obj = {}

    if (element.type === 'tag') {

        const heading = headingForTag(element.name)
        if (heading) {
            type = 'Paragraph'
            obj.heading = heading
        }

        const decoration = decorationForTag(element.name)
        if (decoration) {
            type = 'TextRun'
            obj[decoration] = true
        }

        if (element.parent.name === 'ul' && element.name === 'li') {
            type = 'Paragraph'
            obj.bullet = { level: 0 }
        }

        if (element.children.length > 0) {
            if (type === 'TextRun') {
                for (const child of element.children) {
                    const converted = convertElement(child, true)
                    if (converted) {
                        Object.assign(obj, converted)
                    }
                }
            }
            else {
                obj.children = []
                for (const child of element.children) {
                    const converted = convertElement(child)
                    if (converted) {
                        obj.children.push(converted)
                    }
                }
            }
        }
    }
    else if (element.type === 'text') {
        if (element.data.trim() === "") {
            skip = true
        }
        else {
            type = 'TextRun'
            obj.text = element.data
        }
    }

    if (skip) {
        return false
    }

    if (objectOnly || runType === 'debug') {
        return obj
    }

    if (type === 'Paragraph') {
        return new Paragraph(obj)
    }
    if (type === 'TextRun') {
        return new TextRun(obj)
    }
    //if (element.name === 'ul') {
    //    const para = new Paragraph(obj)
    //    para.bullets()
    //    return para
    //}

    return obj
}

function headingForTag(tag) {
    switch (tag) {
        case 'h1': return HeadingLevel.HEADING_1
        case 'h2': return HeadingLevel.HEADING_2
        case 'h3': return HeadingLevel.HEADING_3
        case 'h4': return HeadingLevel.HEADING_4
        case 'h5': return HeadingLevel.HEADING_5
        case 'h6': return HeadingLevel.HEADING_6
        default: return false
    }
}

function decorationForTag(tag) {
    switch (tag) {
        case 'em': return 'italics'
        case 'strong': return 'bold'
        default: return false
    }
}

function docxFromHtml(html) {
    const doc = new Document()
    doc.addSection({
        properties: {},
        children: [
            new Paragraph({
                text: 'Hello world'
            })
        ]
    })
    return Packer.toBuffer(doc)
}

module.exports = docxFromHtml
