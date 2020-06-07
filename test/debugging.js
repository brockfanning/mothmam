const fs = require('fs')
const { convertToHtml } = require('mammoth')
const docxFromHtml = require('../lib/index')
const snippets = require('./snippets.json')

//docxFromHtml(snippets['List with emphasis']).then(() => {
//    console.log('Finished')
//})