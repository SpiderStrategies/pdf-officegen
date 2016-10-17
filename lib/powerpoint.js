'use strict'

var fs = require('fs')
var os = require('os')

var officegen = require('officegen')
var PDF2Images = require('pdf2images-multiple')
var Promise = require('bluebird')
var rmdir = require('rimraf')

// Loggers
var debug = require('debug')
var errorLogger = debug('app:error')
var pdfLogger = debug('app:pdf')
var debugLogger = debug('debug:pdf')
pdfLogger.log = console.log.bind(console)
debugLogger.log = console.log.bind(console)

class Powerpoint {

  /**
   *
   * @param pdfFile
   * @param pptxOutput
   * @param done
   * @param options
   * @param options.stagingDir - A directory where intermediate PNG images will
   * be placed when converting into slides.  If a file with the same name
   *   exists ImageMagick will not process that image again.  Therefore, it is
   *   recommended that a different folder be used for each conversion.  If
   *   undefined, a random directory will be created under the systems temp
   *   directory.  It will be deleted once the job has completed.
   * @param options.convertOptions - ImageMagick conversion options (minus the -)
   *   Currently supported: density(<300)
   */
  convertPDFToImages (pdfFile, pptxOutput, done, options) {
    pdfLogger('converting', pdfFile)
    const opts = options || {}
    const stagingDir = this._getStagingDirectory(opts.stagingDir)
    stagingDir.then((outputDir) => {
      var pdf2images = this._getImageMagickOptions(pdfFile, this._getConvertOptions(opts), outputDir)
      this._convert(pdf2images, outputDir, pptxOutput, done)
    }, (err) => done(err))
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
    converter.pdf.convert((err, imagePath) => {
      if (err) errorLogger(err)
      // Will it be faster to add images here?
    }, (err, images) => {
      if (err) {
        done(err)
      } else {
        this._aggregateSlides(images, pptxOutput, done, outputDir)
      }
    })
  }

  _aggregateSlides (images, pptxOutput, done, outputDir) {
    pdfLogger('creating slides...')
    this._createSlides(images, pptxOutput, (slideErr, output) => {
      pdfLogger('Finished rendering all slides')
      done(slideErr, output)
      rmdir(outputDir,
        (e) => e && errorLogger('Could not delete working directory:', e))
    })
  }

  _createSlides (imageFiles, pptFile, done) {
    pdfLogger('Adding images to slides')
    var pptx = officegen('pptx')
    this._addSlidesToPresentation(imageFiles, pptx)
    this._savePresentationFile(pptFile, done, pptx)
  }

  _addSlidesToPresentation (imageFiles, pptx) {
    var sortedImages = imageFiles.sort((a, b) => {
      return a.localeCompare(b)
    })
    debugLogger('Sorted Images:', sortedImages)

    sortedImages.forEach(i => {
      var slide = pptx.makeNewSlide()
      slide.addImage(i, {x: 0, y: 0, cx: '100%', cy: '100%'})
    })
  }

  _savePresentationFile (pptFile, done, pptx) {
    var out = fs.createWriteStream(pptFile)
    out.on('close', () => {
      pdfLogger('Created the PPTX file:', pptFile)
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

}

module.exports = Powerpoint
