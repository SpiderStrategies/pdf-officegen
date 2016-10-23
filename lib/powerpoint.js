'use strict'

var EventEmitter2 = require('eventemitter2').EventEmitter2
var fs = require('fs')
var os = require('os')

var officegen = require('officegen')
var PDF2Images = require('pdf2images-multiple')
var Promise = require('bluebird')
var rmdir = require('rimraf')

// Loggers
var debug = require('debug')
var errorLogger = debug('pdfppt:error')
var pdfLogger = debug('pdfppt:app')
var debugLogger = debug('pdfppt:debug')
pdfLogger.log = console.log.bind(console)
debugLogger.log = console.log.bind(console)

class Powerpoint extends EventEmitter2 {

  constructor(){
    super({
      wildcard: true, //Allow clients to listen based on wildcard expressions
      maxListeners: 10 //Node Default, set to 0 for Infinity
    })
  }

  /**
   *
   * @param pdfFile
   * @param pptxOutput
   * @param options
   * @param options.stagingDir - A directory where intermediate PNG images will
   * be placed when converting into slides.  If a file with the same name
   *   exists ImageMagick will not process that image again.  Therefore, it is
   *   recommended that a different folder be used for each conversion.  If
   *   undefined, a random directory will be created under the systems temp
   *   directory.  It will be deleted once the job has completed.
   * @param options.convertOptions - ImageMagick conversion options (minus the -)
   *   Currently supported: density(<300)
   * @param done
   */
  convertPDFToImages (pdfFile, pptxOutput, options, done) {
    pdfLogger('converting', pdfFile)

    let callback
    let opts = {}

    if (typeof arguments[2] == 'object'){
      opts = options
      callback = done
    }else{
      callback = arguments[2]
    }

    pdfLogger('options:', opts)

    const stagingDir = this._getStagingDirectory(opts.stagingDir)
    stagingDir.then((outputDir) => {
      var pdf2images = this._getImageMagickOptions(pdfFile, this._getConvertOptions(opts), outputDir)
      this._convert(pdf2images, outputDir, pptxOutput, callback)
    }, (err) => callback(err))
  }

  _getConvertOptions (opts) {
    const convertOpts = opts.convertOptions || {}
    const convertOptions = {
      '-density': convertOpts.density ? Math.min(convertOpts.density, 300) : 72, // DPI
      // '-resize' : '800x600',
      // '-trim': '',
      // '-sharpen' : '0x1.0'
      '-quality': 100
    }
    return convertOptions
  }

  _getImageMagickOptions (pdfFile, convertOptions, outputDir) {
    pdfLogger('Using staging directory:', outputDir)
    const options = {
      convert_options: convertOptions, // optional
      // convert_operators: convert_operators, //optional
      output_dir: outputDir, // optional
      ext: 'png', // optional, png is the default value
      gm: false // Use GraphicksMagic //optional, false is the default value
    }
    pdfLogger('ImageMagick Options:', options)
    return PDF2Images(pdfFile, options)
  }

  _convert (converter, outputDir, pptxOutput, done) {
    pdfLogger('converting', pptxOutput, 'into images...')
    var self = this
    var start = this.nowInMillis()
    var imgStart = start
    converter.pdf.convert((err, imagePath) => {
      if (err) {
        self.emit('err.png.single', {output: imagePath, error: err, time: this.elapsed(imgStart)})
        errorLogger(err)
      }else{
        self.emit('done.png.single', {output: imagePath, time: this.elapsed(imgStart)})
      }
      imgStart = this.nowInMillis()
      // Will it be faster to add images here?
    }, (err, images) => {
      if (err) {
        done(err)
        self.emit('err.png.all', {output: images, error: err, time: this.elapsed(start)})
      } else {
        self.emit('done.png.all', {output: images, time: this.elapsed(start)})
        this._aggregateSlides(images, pptxOutput, done, outputDir)
      }
    })
  }

  _aggregateSlides (images, pptxOutput, done, outputDir) {
    this._createSlides(images, pptxOutput, (slideErr, output) => {
      pdfLogger('Finished rendering all slides')
      done(slideErr, output)
      rmdir(outputDir,
        (e) => e && errorLogger('Could not delete working directory:', e))
    })
  }

  _createSlides (imageFiles, pptFile, done) {
    pdfLogger('Adding images to slides')
    var start = this.nowInMillis()
    var pptx = officegen('pptx')
    this._addSlidesToPresentation(imageFiles, pptx)
    this.emit('done.pptx.creation', {output: pptFile, time: this.elapsed(start)})
    this._savePresentationFile(pptFile, done, pptx)
  }

  _addSlidesToPresentation (imageFiles, pptx) {
    // Images are not in order, so sort them by name
    var sortedImages = imageFiles.sort((a, b) => {
      // TODO: test with more than 10 images (string comparison)
      return a.localeCompare(b)
    })
    debugLogger('Sorted Images:', sortedImages)

    //TODO: Need a callback here if this blocks too long
    sortedImages.forEach(i => {
      var slide = pptx.makeNewSlide()
      slide.addImage(i, {x: 0, y: 0, cx: '100%', cy: '100%'})
    })
  }

  _savePresentationFile (pptFile, done, pptx) {
    var out = fs.createWriteStream(pptFile)
    out.on('close', () => {
      pdfLogger('Created the PPTX file:', pptFile)
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
        resolve(folder)
      })
    })
  }

  nowInMillis() {
    return Date.now() //process.hrtime()[1] / 1000000
  }

  elapsed (start) {
    return this.nowInMillis() - start
  }

}

module.exports = Powerpoint
