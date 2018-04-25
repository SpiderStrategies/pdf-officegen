const _ = require('lodash')
const officegen = require('spider-officegen')

const OfficeDoc = require('./office-doc')
const {nowInMillis, elapsed} = require('./util')

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
    width: 624,
    height: 864
  }
}

class Word extends OfficeDoc {

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
    docx.on('error', function (err) {
      console.log(err)
    })
    const imgDim = this.options.imageDimensions
    imageFiles.forEach(i => {
      var pObj = docx.createP()
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
