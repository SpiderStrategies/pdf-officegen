'use strict'

const EventEmitter2 = require('eventemitter2').EventEmitter2

const fs = require('fs')
const path = require('path')

const _ = require('lodash')
const officegen = require('officegen')
const rmdir = require('rimraf')

const Engine = require('./engine')
const util = require('./util')
const {info} = require('./logger')

const defaultOptions = {
  clean: true,
  cropLastImage: false,
  dimensions: { width: 800, height: 600, type: 'screen4x3' },
  jobId: ''
}

class Powerpoint extends EventEmitter2 {

  /**
   *
   * @param options
   * @param options.clean {boolean=true} set to false if intermediate image
   *   files should be left on the filesystem.
   * @param options.jobId {String} if provided, this will be included in any
   *   logging output
   * @param {boolean} [options.cropLastImage=false] requires ImageMagick
   *   `convert` to be on the path.  Will crop the last pdf image before
   *   placing on slide.
   * @param {Number} [options.dimensions.width=800] of slides in pixels
   * @param {Number} [options.dimensions.height=600] of slides in pixels
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
   * @param options
   * @param options.stagingDir - A directory where intermediate PNG images will
   *   be placed when converting into slides.  A different folder should be
   *   used for each conversion.  If undefined, a random directory will be
   *   created under the systems temp directory.  It will be deleted once the
   *   job has completed.
   *
   * @param options.convertOptions - ImageMagick conversion options (minus the
   *   -) Currently supported: density(<300)
   * @param done
   */
  convertPDFToPowerpoint (pdfFiles, options, done) {
    let callback
    let opts = {}

    // TODO: Test this
    if (typeof arguments[1] === 'object') {
      opts = options
      callback = done
    } else {
      callback = arguments[1]
    }

    // info('options:', opts)

    util.getStagingDirectory(opts.stagingDir)
      .then((outputDir) => {
        this.imgDir = path.resolve(outputDir, 'img')
        this.pdfDir = path.resolve(outputDir, 'pdf')
        const {engine} = opts
        let convertPromise = new Engine({engine}).convert(outputDir, pdfFiles)
        convertPromise.then(sortedImages => {
          this._createPowerpoint(outputDir, sortedImages, callback)
        })
      }, (err) => callback(err))
      .catch(err => callback(err))
  }

  _createPowerpoint (outputDir, sortedImages, callback) {
    const pptxOutput = path.resolve(outputDir, `output_${process.hrtime()[1]}.pptx`)
    this._aggregateSlides(sortedImages, pptxOutput, this.imgDir, callback)
  }

  /**
   *
   * @param {Array} images
   * @param pptxOutput pptx file path
   * @param imgDir the directory where png files are generated
   * @param done callback
   * @private
   */
  _aggregateSlides (images, pptxOutput, imgDir, done) {
    this._createSlides(images, pptxOutput, (slideErr, output) => {
      done(slideErr, output)
      if (this.options.clean) {
        const start = this.nowInMillis()
        rmdir(imgDir, (err) => {
          if (err) {
            this.emit('done.png.clean', {output: imgDir, time: this.elapsed(start), error: err})
            info(this.options.jobId, 'Could not delete working directory:', imgDir, err)
          }
        })
      }
    })
  }

  _createSlides (imageFiles, pptFile, done) {
    const start = this.nowInMillis()
    const pptx = officegen('pptx')
    const d = this.options.dimensions
    // https://github.com/Ziv-Barber/officegen/issues/112
    pptx.setSlideSize(d.width, d.height, d.type)
    this._addSlidesToPresentation(imageFiles, pptx)
    this.emit('done.pptx.creation', {time: this.elapsed(start)})
    this._savePresentationFile(pptFile, done, pptx)
  }

  _addSlidesToPresentation (imageFiles, pptx) {
    // TODO: Need a callback here if this blocks too long
    imageFiles.forEach(i => { pptx.makeNewSlide().addImage(i) })
  }

  _savePresentationFile (pptFile, done, pptx) {
    const start = this.nowInMillis()
    const out = fs.createWriteStream(pptFile)
    out.on('close', () => {
      this.emit('done.pptx.saved', {output: pptFile, time: this.elapsed(start)})
      done(null, pptFile)
    })
    pptx.generate(out)
  }

}

module.exports = Powerpoint
