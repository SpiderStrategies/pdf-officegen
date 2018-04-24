const _ = require('lodash')
const officegen = require('officegen')

const OfficeDoc = require('./office-doc')
const {nowInMillis, elapsed} = require('./util')
const {info} = require('./logger')

class Powerpoint extends OfficeDoc {

  constructor (options) {
    super(_.extend({}, options, {extension: 'pptx'}))
  }

  createDocument (imageFiles, pptFile, done) {
    const start = nowInMillis()
    const pptx = officegen('pptx')
    const d = this.options.dimensions
    info('slide dimensions', d)
    // https://github.com/Ziv-Barber/officegen/issues/112
    pptx.setSlideSize(d.width, d.height, d.type)
    imageFiles.forEach(i => { pptx.makeNewSlide().addImage(i) })
    this.emit('done.pptx.creation', {time: elapsed(start)})
    this.saveOfficeDocument(pptFile, done, pptx)
  }
}

module.exports = Powerpoint
