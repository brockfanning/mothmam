# Mothmam

Convert HTML code to simple simple Word documents.

## Mammoth inverted

This library is intended to be the inverse (so to speak) of the wonderful
[Mammoth.js](https://github.com/mwilliamson/mammoth.js) library. Whereas Mammoth
converts Word files into simple HTML, this library converts HTML into simple Word
files.

## Usage

```
const fs = require('fs')
const { convertToWord } = require('mothmam')

const html = '<p>Hello world</p>'
convertToWord(html).then((docx) => {
    fs.writeFileSync('example.docx', docx)
})
```
