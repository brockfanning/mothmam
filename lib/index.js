const cheerio = require('cheerio')
const { Paragraph, TextRun, Document, Packer, AlignmentType, HeadingLevel} = require('docx')


function docxFromHtml(html) {
    const $ = cheerio.load(html)
    const doc = new Document(getDocumentProperties())
    const paragraphs = []

    $('body').contents('*').each((idx, element) => {
        convertElement(element).forEach(item => {
            paragraphs.push(item)
        })
    })

    doc.addSection({
        properties: {},
        children: paragraphs
    })
    return Packer.toBuffer(doc)
}


/**
 * Returns an array of objects for use by Docx.js.
 */
function convertElement(element, objectOnly=false) {

    // Short names for properties of the element.
    const { children, type, data, name } = element

    // Internal stuff to keep track of multiple items to return.
    const items = []
    function addItem() {
        items.push({})
    }
    function getItem() {
        return items[items.length - 1]
    }

    // Skip empty text nodes.
    if (type === 'text' && data.trim() === "") {
        return []
    }

    // Start us off with a single item.
    addItem()

    if (type === 'tag') {
        if (tagIsList(name)) {
            return convertList(element)
        }
        if (tagIsParagraph(name)) {
            getItem().docxClass = 'Paragraph'
        }
        if (tagIsHeading(name)) {
            getItem().heading = headingForTag(name)
        }
        if (tagIsText(name)) {
            getItem().docxClass = 'TextRun'
        }
        if (tagIsEmphasis(name)) {
            getItem()[emphasisForTag(name)] = true
        }
    }
    else if (type === 'text') {
        getItem().docxClass = 'TextRun'
        getItem().text = data
    }

    if (children && children.length > 0) {
        if (getItem().docxClass === 'TextRun') {
            children.forEach(child => {
                convertElement(child, true).forEach(converted => {
                    Object.assign(getItem(), converted)
                })
            })
        }
        else {
            getItem().children = []
            children.forEach(child => {
                convertElement(child, objectOnly).forEach(converted => {
                    getItem().children.push(converted)
                })
            })
        }
    }

    const returnConverted = []
    items.forEach(item => {
        if (Object.keys(item) === 0) {
            return
        }
        if (objectOnly) {
            returnConverted.push(item)
        }
        else if (item.converted) {
            returnConverted.push(item.converted)
        }
        else if (item.docxClass === 'Paragraph') {
            returnConverted.push(new Paragraph(item))
        }
        else if (item.docxClass === 'TextRun') {
            returnConverted.push(new TextRun(item))
        }
    })
    return returnConverted
}

function convertList(list, listLevel=0) {
    const flatList = []
    list.children.forEach(child => {
        convertListItem(child, listLevel).forEach(converted => {
            flatList.push(converted)
        })
    })
    return flatList
}

function convertListItem(listItem, listLevel) {
    const nestedChildren = []
    const nonNestedChildren = []
    listItem.children.forEach(child => {
        if (child.type === 'text') {
            nonNestedChildren.push(new TextRun(child.data))
        }
        else if (child.type === 'tag') {
            if (tagIsText(child.name) || tagIsParagraph(child.name)) {
                nonNestedChildren.push(convertTextualElement(child))
            }
            else if (tagIsList(child.name)) {
                convertList(child, listLevel + 1).forEach(converted => {
                    nestedChildren.push(converted)
                })
            }
        }
    })
    const flatList = []
    flatList.push(new Paragraph(Object.assign({
        children: nonNestedChildren,
    }, getListOptions(listItem, listLevel))))
    nestedChildren.forEach(child => {
        flatList.push(child)
    })

    return flatList
}

function getListOptions(listItem, listLevel) {
    if (listItem.parent.name === 'ul') {
        return {
            bullet: {
                level: listLevel,
            }
        }
    }
    else {
        return {
            numbering: {
                reference: 'decimal-numbering',
                level: listLevel,
            }
        }
    }
}

function convertTextualElement(element) {
    const obj = {}
    if (element.type === 'text') {
        obj.text = element.data
    }
    else if (element.type === 'tag') {
        if (tagIsEmphasis(element.name)) {
            obj[emphasisForTag(element.name)] = true
        }
        if (element.children.length) {
            obj.children = []
            element.children.forEach(child => {
                const converted = convertTextualElement(child)
                if (converted) {
                    obj.children.push(converted)
                }
            })
        }
    }
    return new TextRun(obj)
}

function tagIsParagraph(tag) {
    return ['p'].includes(tag) || tagIsHeading(tag)
}

function tagIsHeading(tag) {
    return ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'].includes(tag)
}

function tagIsList(tag) {
    return ['ul', 'ol'].includes(tag)
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

function getDocumentProperties() {
    return {
        numbering: {
            config: [
                {
                    levels: [
                        {
                            level: 0,
                            format: 'decimal',
                            text: '%1',
                            alignment: AlignmentType.START,
                            style: {
                                paragraph: {
                                    indent: { left: 720, hanging: 260 },
                                },
                            },
                        },
                        {
                            level: 1,
                            format: 'decimal',
                            text: '%1',
                            alignment: AlignmentType.START,
                            style: {
                                paragraph: {
                                    indent: { left: 720, hanging: 260 },
                                },
                            },
                        },
                        {
                            level: 2,
                            format: 'decimal',
                            text: '%1',
                            alignment: AlignmentType.START,
                            style: {
                                paragraph: {
                                    indent: { left: 720, hanging: 260 },
                                },
                            },
                        },
                    ],
                    reference: 'decimal-numbering',
                },
            ],
        },
    }
}

module.exports = docxFromHtml
