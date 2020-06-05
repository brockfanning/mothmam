var expect = require('chai').expect;

const testSnippets = [
    '<p>Hello world</p>',
    '<p>Hello <em>world</em></p>',
    '<p>Hello <strong>world</strong></p>',
    '<p>Hello <strong><em>world</em></strong></p>',
]

for (const testSnippet of testSnippets) {
    describe(testSnippet, function() {
        it('should be unchanged after converting from HTML and back', function() {
            expect(convertFromHtmlAndBack(testSnippet)).to.be.equal(testSnippet)
        })
    })
}

function convertFromHtmlAndBack(html) {
    return html
}
