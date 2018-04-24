const _ = require('lodash')
const EventEmitter2 = require('eventemitter2').EventEmitter2
const rmdir = require('rimraf')

const fs = require('fs')
const path = require('path')

const Engine = require('./engine')
const {info} = require('./logger')
const {nowInMillis, elapsed, getStagingDirectory} = require('./util')

const defaultOptions = {
  clean: true,
  cropLastImage: false,
  dimensions: { width: 800, height: 600, type: 'screen4x3' },
  jobId: ''
}

class OfficeDoc extends EventEmitter2 {

  /**
   *
   * @param {String} options.extension  'pptx', 'docx'
   *
   * @param {Boolean} [options.clean=true] set to false if intermediate image
   *   files should be left on the filesystem.
   *
   * @param {String} [options.jobId] if provided this will be included in any logging output
   *
   * @param {Boolean} [options.cropLastImage=false] requires ImageMagick
   *   `convert` to be on the path.  Will crop the last pdf image before
   *   placing on slide.
   *
   * @param {Number} [options.dimensions.width=800] of slides in pixels
   *
   * @param {Number} [options.dimensions.height=600] of slides in pixels
   *
   * @param {String} [options.dimensions.type=screen4x3] '35mm' 'A3' 'A4'
   *   'B4ISO' 'B4JIS' 'B5ISO' 'B5JIS' 'banner' 'custom' 'hagakiCard' 'ledger'
   *   'letter' 'overhead' 'screen16x10' 'screen16x9' 'screen4x3'
   *
   */
  constructor (options) {
    super({
      wildcard: true, // Allow clients to listen based on wildcard expressions
      maxListeners: 10 // Node Default, set to 0 for Infinity
    })
    this.options = _.merge({}, defaultOptions, options)
    process.nextTick(() => {
      this.emit('options', { output: JSON.stringify(this.options) })
    })
  }

  /**
   *
   * @param pdfFiles {array|string} An array of PDF files that should be
   *   converted
   *
   * @param options.stagingDir - A directory where intermediate PNG images will
   *   be placed when converting into slides.  A different folder should be
   *   used for each conversion.  If undefined, a random directory will be
   *   created under the systems temp directory.  It will be deleted once the
   *   job has completed.
   *
   * @param options.convertOptions - ImageMagick conversion options (minus the
   *   -) Currently supported: density(<300)
   *
   * @param done
   */
  convertFromPdf (pdfFiles, options, done) {
    let callback
    let opts = {}

    if (typeof arguments[1] === 'object') {
      opts = options
      callback = done
    } else {
      callback = arguments[1]
    }

    getStagingDirectory(opts.stagingDir)
        .then((outputDir) => {
          const imgDir = path.resolve(outputDir, 'img')
          const pdfDir = path.resolve(outputDir, 'pdf')
          const {engine, extension, cropLastImage, convertOptions} = this.options

          const engineOpts = {engine, cropLastImage, convertOptions, imgDir, pdfDir}
          const conversionEngine = new Engine(engineOpts)
          conversionEngine.onAny((name, result) => this.emit(name, result))
          conversionEngine.convert(outputDir, pdfFiles).then(sortedImages => {
            const filename = `output_${process.hrtime()[1]}.${extension}`
            const outputFile = path.resolve(outputDir, filename)
            this._aggregatePageImages(sortedImages, outputFile, imgDir, callback)
          })
        }, (err) => callback(err))
        .catch(err => callback(err))
  }

  /**
   *
   * @param {Array} images
   * @param outputFile output file path
   * @param imgDir the directory where png files are generated
   * @param done callback
   * @private
   */
  _aggregatePageImages (images, outputFile, imgDir, done) {
    this.createDocument(images, outputFile, (slideErr, output) => {
      done(slideErr, output)
      if (this.options.clean) {
        const start = nowInMillis()
        rmdir(imgDir, (err) => {
          if (err) {
            this.emit('done.png.clean', {output: imgDir, time: elapsed(start), error: err})
            info(this.options.jobId, 'Could not delete working directory:', imgDir, err)
          }
        })
      }
    })
  }

  saveOfficeDocument (file, done, officeGenDoc) {
    const start = nowInMillis()
    const out = fs.createWriteStream(file)
    out.on('close', () => {
      this.emit(`done.${this.options.extension}.saved`, {output: file, time: elapsed(start)})
      done(null, file)
    })
    officeGenDoc.generate(out)
  }
}

module.exports = OfficeDoc
