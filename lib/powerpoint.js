const _ = require('lodash')
const officegen = require('spider-officegen')

const OfficeDoc = require('./office-doc')
const {nowInMillis, elapsed} = require('./util')

const defaultOptions = {
  extension: 'pptx',
  clean: true,
  cropLastImage: false,
  dimensions: { width: 800, height: 600, type: 'screen4x3' },
  jobId: ''
}

class Powerpoint extends OfficeDoc {

  /**
   * @param {Number} [options.dimensions.width=800] of slides in pixels
   *
   * @param {Number} [options.dimensions.height=600] of slides in pixels
   *
   * @param {String} [options.dimensions.type=screen4x3] '35mm' 'A3' 'A4'
   *   'B4ISO' 'B4JIS' 'B5ISO' 'B5JIS' 'banner' 'custom' 'hagakiCard' 'ledger'
   *   'letter' 'overhead' 'screen16x10' 'screen16x9' 'screen4x3'
   *
   * @constructs
   */
  constructor (options) {
    super(_.merge({}, defaultOptions, options))
  }

  createDocument (imageFiles, pptFile, done) {
    const start = nowInMillis()
    const pptx = officegen('pptx')
    pptx.on('error', error => {
      this.emit('err.pptx', {error})
    })
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
