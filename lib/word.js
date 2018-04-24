const _ = require('lodash')
const officegen = require('officegen')

const OfficeDoc = require('./office-doc')
const {nowInMillis, elapsed} = require('./util')

class Word extends OfficeDoc {

  constructor (options) {
    super(_.extend({}, options, {extension: 'docx'}))
  }

  createDocument (imageFiles, wordFile, done) {
    const start = nowInMillis()
    const docx = officegen('docx')
    imageFiles.forEach(i => {
      var pObj = docx.createP()
      pObj.addImage(i)
      docx.putPageBreak()
    })
    this.emit('done.docx.creation', {time: elapsed(start)})
    this.saveOfficeDocument(wordFile, done, docx)
  }
}

module.exports = Word
