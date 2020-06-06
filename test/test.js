const fs = require('fs')
const { expect } = require('chai')
const { convertToHtml } = require('mammoth')
const docxFromHtml = require('../lib/index')

const testSnippets = {
    'Paragraph tag': '<p>Hello world</p>',
    'H1 and paragraph tags': '<h1>Hello world</h1><p>Foobar</p>',
    'Paragraph tag with emphasis': '<p>Hello <em>world</em></p>',
    'Paragraph tag with bold': '<p>Hello <strong>world</strong></p>',
    'Paragraph tag with emphasis and bold': '<p>Hello <strong><em>world</em></strong></p>',
    'Unordered list': '<ul><li>Hello</li><li>world</li></ul>',
    'Ordered list': '<ol><li>Hello</li><li>world</li></ol>',
    'List with bold': '<ul><li><strong>world</strong></li></ul>',
    'List with emphasis': '<ul><li><em>world</em></li></ul>',
    'List with list': '<ul><li>Fruits:<ul><li>Orange</li><li>Apple</li></ul></li><li>Veggies<ul><li>Ocra</li><li>Peas</li></ul></li></ul>',
    'List with list with emphasis': '<ul><li>Fruits:<ul><li><em>the</em> orange</li></ul></li></ul>',
}

for (const [description, snippet] of Object.entries(testSnippets)) {
    describe(description, function() {
        it('should be unchanged after converting from HTML and back', async () => {
            const wordFromHtml = await docxFromHtml(snippet)
            fs.writeFileSync(description + '.docx', wordFromHtml)
            const htmlFromWord = await convertToHtml({ buffer: wordFromHtml})
            expect(htmlFromWord.value).to.be.equal(snippet)
        })
    })
}
