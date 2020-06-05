const cheerio = require('cheerio')
const { Paragraph, TextRun, Document, Packer, HeadingLevel} = require('docx')

function convertElement(element, objectOnly=false) {

    let skip = false
    let type = null
    const obj = {}

    if (element.type === 'tag') {

        if (tagIsParagraph(element.name)) {
            type = 'Paragraph'
        }

        if (tagIsHeading(element.name)) {
            obj.heading = headingForTag(element.name)
        }

        if (tagIsText(element.name)) {
            type = 'TextRun'
        }

        if (tagIsEmphasis(element.name)) {
            obj[emphasisForTag(element.name)] = true
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

    if (objectOnly) {
        return obj
    }

    if (type === 'Paragraph') {
        return new Paragraph(obj)
    }
    if (type === 'TextRun') {
        return new TextRun(obj)
    }

    return obj
}

function tagIsParagraph(tag) {
    return ['p'].includes(tag) || tagIsHeading(tag)
}

function tagIsHeading(tag) {
    return ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'].includes(tag)
}

function tagIsText(tag) {
    return ['span'].includes(tag) || tagIsEmphasis(tag)
}

function tagIsEmphasis(tag) {
    return ['em', 'strong'].includes(tag)
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

function emphasisForTag(tag) {
    switch (tag) {
        case 'em': return 'italics'
        case 'strong': return 'bold'
        default: return false
    }
}

function docxFromHtml(html) {
    const $ = cheerio.load(html)
    const doc = new Document()
    const sections = []

    $('body').contents('*').each((idx, element) => {
        const converted = convertElement(element)
        if (converted) {
            sections.push(converted)
        }
    })

    doc.addSection({
        properties: {},
        children: sections
    })
    return Packer.toBuffer(doc)
}

module.exports = docxFromHtml
