const mammoth = require('mammoth')

mammoth.convertToHtml({path: 'example.docx' })
    .then(function(result) {
        console.log(result.value)
        console.log(result.messages)
    })
    .done();
