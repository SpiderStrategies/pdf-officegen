const _ = require('lodash')
const officegen = require('spider-officegen')

const OfficeDoc = require('./office-doc')
const {nowInMillis, elapsed} = require('./util')

/**
 * Inspected document with 1" margins to find this value
 * @type {number}
 */
const ONE_INCH_MARGIN = 1440

const defaultOptions = {
  extension: 'docx',
  clean: true,
  cropLastImage: false,
  jobId: '',
  orientation: 'portrait',
  pageMargins: {
    top: ONE_INCH_MARGIN,
    left: ONE_INCH_MARGIN,
    bottom: ONE_INCH_MARGIN,
    right: ONE_INCH_MARGIN
  },
  imageDimensions: {
    width: 624, // 6.5" x 96 pixels
    height: 864 // 9" x 96 pixels
  }
}

class Word extends OfficeDoc {
  /**
   * @param {Object} [options.pageMargins] 1440 = 1 inch
   * @param {Number} [options.pageMargins.top=1440] top page margin
   * @param {Number} [options.pageMargins.bottom=1440] bottom page margin
   * @param {Number} [options.pageMargins.left=1440] left page margin
   * @param {Number} [options.pageMargins.right=1440] right page margin
   *
   * @param {Number} [options.imageDimensions.width=624] in pixels;
   * should be the page width minus left/right margin
   *
   * @param {Number} [options.imageDimensions.height=864] in pixels;
   * should be the page height minus top/bottom margin
   *
   * @param {String} [options.orientation='portrait'] 'portrait' and 'landscape'
   * are supported by officegen.
   *
   * @constructs
   */
  constructor (options = {}) {
    super(_.merge({}, defaultOptions, options))
  }

  createDocument (imageFiles, wordFile, done) {
    // officegen.setVerboseMode(true)
    const start = nowInMillis()
    const docx = officegen({
      type: 'docx',
      orientation: this.options.orientation,
      pageMargins: this.options.pageMargins
    })
    docx.on('error', error => {
      this.emit('err.docx', {error})
    })
    const imgDim = this.options.imageDimensions
    imageFiles.forEach(i => {
      var pObj = docx.createP()
      // Have to specify the dimensions for higher resolution images (e.g. 150 dpi)
      // that are scaled down to fit on a single page
      const options = {
        cx: imgDim.width,
        cy: imgDim.height
      }
      pObj.addImage(i, options)
    })
    this.emit('done.docx.creation', {time: elapsed(start)})
    this.saveOfficeDocument(wordFile, done, docx)
  }
}

module.exports = Word
