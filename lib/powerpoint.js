'use strict'

var EventEmitter2 = require('eventemitter2').EventEmitter2
var exec = require('child_process').exec
var fs = require('fs')
var os = require('os')
var path = require('path')

const _ = require('lodash')
var officegen = require('officegen')
var Promise = require('bluebird')
var rmdir = require('rimraf')

// Loggers
var debug = require('debug')
var errorLogger = debug('pdfppt:error')
var pdfLogger = debug('pdfppt:app')
var debugLogger = debug('pdfppt:debug')
pdfLogger.log = console.log.bind(console)
debugLogger.log = console.log.bind(console)

/** allows the caller to provide a path to the GhostScript Executable */
const GS_PATH = process.env.PDF_PPT_GSPATH

/** Going above 300 has a significant impact on performance
 * without much noticeable quality improvement */
const DPI_DENSITY_MAX = 300
/** 72 is the default (because of long-time standard) */
const DPI_DENSITY_DEFAULT = 72

class Powerpoint extends EventEmitter2 {

  /**
   *
   * @param options
   * @param options.clean {boolean=true} set to false if intermediate image files
   * should be left on the filesystem.
   *
   */
  constructor (options) {
    super({
      wildcard: true, // Allow clients to listen based on wildcard expressions
      maxListeners: 10 // Node Default, set to 0 for Infinity
    })
    this.options = options || {}
    this.clean = this.options.clean || true
  }

  /**
   *
   * @param pdfFiles {array|string} An array of PDF files that should be converted
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

    // pdfLogger('options:', opts)

    const stagingDir = this._getStagingDirectory(opts.stagingDir)
    stagingDir.then((outputDir) => {
      this._convertWithGhostScript(outputDir, pdfFiles, options, callback)
    }, (err) => callback(err))
  }

  /**
   * GhostScript can be invoked directly, since ImageMagick just delegates to it
   *
   * @param outputDir
   * @param pdfFiles
   * @param options
   * @param callback
   * @private
   */
  _convertWithGhostScript (outputDir, pdfFiles, options, callback) {
    var start = this.nowInMillis()

    const gsExecutable = os.platform() === 'win32' ? GS_PATH || 'gswin32c.exe' : 'gs'
    const imgDir = path.resolve(outputDir, 'img')
    const co = this._getConvertOptions(options)
    const gsCmdRoot = `"${gsExecutable}" -q -dQUIET -dSAFER -sDEVICE=pngalpha -dMaxBitmap=500000000 -r${co.density} -dUseArtBox`

    // Get the image files for each PDF
    let gsErr = []

    let requests = pdfFiles.map((pdfPath, pdfIndex) => {
      return new Promise((resolve) => {
        const gsCmd = gsCmdRoot + ` -o ${imgDir}/img-${pdfIndex}-%d.png ${pdfPath}`
        exec(gsCmd, (err, stdout, stderr) => {
          this.emit('done.gs.convert', { command: gsCmd, time: this.elapsed(start), error: err })
          if (err) {
            gsErr.push(err)
          }
          resolve()
        })
      })
    })

    // GS executes each PDF asynchronously, so we need a collection of promises
    // to wait for all image files to be present
    Promise.all(requests).then(() => {
      if (!_.isEmpty(gsErr)) {
        this.emit('err.png.all', {error: gsErr, time: this.elapsed(start)})
      } else {
        const images = fs.readdirSync(imgDir).map(f => `${imgDir}/${f}`)
        this.emit('done.png.all', {output: images, time: this.elapsed(start)})
        var pptxOutput = path.resolve(outputDir, `output_${process.hrtime()[1]}.pptx`)
        this._aggregateSlides(images, pptxOutput, callback, outputDir)
      }
    })
  }

  _getConvertOptions (options) {
    const co = options.convertOptions || {}
    const o = {}
    /* Note: if the density is too low and there is a slide with a transparent background,
     The image may show a horizontal line on the slide when it is rendered in the PPTX.
     (was visible at 72, but not visible at 150)
     */
    o.density = co.density ? Math.min(co.density, DPI_DENSITY_MAX) : DPI_DENSITY_DEFAULT
    return o
  }

  /**
   *
   * @param {array} images
   * @param pptxOutput pptx file path
   * @param done callback
   * @param outputDir the directory where png and pptx files are generated
   * @private
   */
  _aggregateSlides (images, pptxOutput, done, outputDir) {
    this._createSlides(images, pptxOutput, (slideErr, output) => {
      done(slideErr, output)
      this.clean && rmdir(path.resolve(outputDir, 'img'), (e) => e && errorLogger('Could not delete working directory:', e))
    })
  }

  _createSlides (imageFiles, pptFile, done) {
    var start = this.nowInMillis()
    var pptx = officegen('pptx')
    this._addSlidesToPresentation(imageFiles, pptx)
    this.emit('done.pptx.creation', {time: this.elapsed(start)})
    this._savePresentationFile(pptFile, done, pptx)
  }

  _addSlidesToPresentation (imageFiles, pptx) {
    // Images are not in order, so sort them by name
    var sortedImages = imageFiles.sort((a, b) => {
      // TODO: test with more than 10 images (string comparison)
      return a.localeCompare(b)
    })
    debugLogger('Sorted Images:', sortedImages)

    // TODO: Need a callback here if this blocks too long
    sortedImages.forEach(i => {
      var slide = pptx.makeNewSlide()
      slide.addImage(i, {x: 0, y: 0, cx: '100%', cy: '100%'})
    })
  }

  _savePresentationFile (pptFile, done, pptx) {
    var out = fs.createWriteStream(pptFile)
    out.on('close', () => {
      this.emit('done.pptx.saved', {output: pptFile})
      done(null, pptFile)
    })
    pptx.generate(out)
  }

  _getStagingDirectory (stagingDir) {
    return new Promise((resolve, reject) => {
      if (stagingDir) {
        fs.stat(stagingDir, (err, s) => {
          if (err || !s.exists()) {
            pdfLogger('staging directory:', stagingDir, 'does not exist, creating a new one')
            return this._createTempStagingDirectory()
          }
          resolve(stagingDir)
        })
      } else {
        resolve(this._createTempStagingDirectory())
      }
    })
  }

  _createTempStagingDirectory () {
    return new Promise((resolve, reject) => {
      fs.mkdtemp(`${os.tmpdir()}/pdf_ppt_`, (err, folder) => {
        if (err) reject(err)
        fs.mkdir(path.resolve(folder, 'img'), (err) => {
          if (err) reject(err)
          resolve(folder)
        })
      })
    })
  }

  nowInMillis () {
    return Date.now() // process.hrtime()[1] / 1000000
  }

  elapsed (start) {
    return this.nowInMillis() - start
  }
}

module.exports = Powerpoint
