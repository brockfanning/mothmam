const fs = require('fs')
const { expect } = require('chai')
const { convertToHtml } = require('mammoth')
const docxFromHtml = require('../lib/index')
const testSnippets = require('./snippets.json')

let tags = null
//tags = ['emphasis', 'list']

for (const testSnippet of testSnippets) {
    if (tags && tags.some(tag => { return !testSnippet.tags.includes(tag) })) continue
    describe(testSnippet.description, function() {
        it('should be unchanged after converting from HTML and back', async () => {
            const wordFromHtml = await docxFromHtml(testSnippet.snippet)
            fs.writeFileSync('TEST - ' + testSnippet.description + '.docx', wordFromHtml)
            const htmlFromWord = await convertToHtml({ buffer: wordFromHtml})
            expect(htmlFromWord.value).to.be.equal(testSnippet.snippet)
        })
    })
}
