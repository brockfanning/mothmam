const { expect } = require('chai')
const { convertToHtml } = require('mammoth')
const docxFromHtml = require('../lib/index')

const testSnippets = [
    '<p>Hello world</p>',
    '<p>Hello <em>world</em></p>',
    '<p>Hello <strong>world</strong></p>',
    '<p>Hello <strong><em>world</em></strong></p>',
]

for (const testSnippet of testSnippets) {
    describe(testSnippet, function() {
        it('should be unchanged after converting from HTML and back', async () => {
            const wordFromHtml = await docxFromHtml(testSnippet)
            const htmlFromWord = await convertToHtml({ buffer: wordFromHtml})
            expect(htmlFromWord.value).to.be.equal(testSnippet)
        })
    })
}
