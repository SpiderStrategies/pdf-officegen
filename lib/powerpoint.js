const _ = require('lodash')
const officegen = require('officegen')

const OfficeDoc = require('./office-doc')
const {nowInMillis, elapsed} = require('./util')

class Powerpoint extends OfficeDoc {

  constructor (options) {
    super(_.extend({}, options, {extension: 'pptx'}))
  }

  createDocument (imageFiles, pptFile, done) {
    const start = nowInMillis()
    const pptx = officegen('pptx')
    const d = this.options.dimensions
    // https://github.com/Ziv-Barber/officegen/issues/112
    pptx.setSlideSize(d.width, d.height, d.type)
    this._addSlidesToPresentation(imageFiles, pptx)
    this.emit('done.pptx.creation', {time: elapsed(start)})
    this.saveOfficeDocument(pptFile, done, pptx)
  }

  _addSlidesToPresentation (imageFiles, pptx) {
    imageFiles.forEach(i => {
      var slide = pptx.makeNewSlide()
      slide.addImage(i, {x: 0, y: 0, cx: '100%', cy: '100%'})
    })
  }
}

module.exports = Powerpoint
