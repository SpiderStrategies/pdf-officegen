'use strict'

var fs = require('fs')
var os = require('os')

var officegen = require('officegen')
var PDF2Images = require('pdf2images-multiple')
var Promise = require("bluebird")
var rmdir = require('rimraf')

// Loggers
var debug = require('debug')
var errorLogger = debug('app:error')
var pdfLogger = debug('app:pdf')
pdfLogger.log = console.log.bind(console)

class Powerpoint {

  /**
   *
   * @param pdfFile
   * @param pptxOutput
   * @param done
   * @param options
   * @param options.stagingDir - A directory where intermediate PNG images will
   * be placed when converting into slides.  If a file with the same name exists
   * ImageMagick will not process that image again.  Therefore, it is recommended
   * that a different folder be used for each conversion.  If undefined, a random
   * directory will be created under the systems temp directory.  It will be deleted
   * once the job has completed.
   */
  convertPDFToImages (pdfFile, pptxOutput, done, options) {
    var convert_options = {
      '-density': 600, //DPI
      //'-resize' : '800x600',
      '-quality': 100
    }
    var convert_operators = {
      //'-trim': '',
      //'-sharpen' : '0x1.0'
    }

    const opts = options || {}
    const stagingDir = this._getStagingDirectory(opts.stagingDir)

    stagingDir.then( (outputDir) => {
      var pdf2images = PDF2Images(pdfFile, {
        convert_options: convert_options, //optional
        // convert_operators: convert_operators, //optional
        output_dir: outputDir, //optional
        ext: 'png', //optional, png is the default value
        gm: false //Use GraphicksMagic //optional, false is the default value
      })

      pdfLogger('converting', pdfFile, 'into images...')

      pdf2images.pdf.convert( (err, image_path) => {
        // Will it be faster to add images here?
      }, (err, images) => {
        if (err) {
          done(err)
        } else {
          pdfLogger('creating slides...')
          this.createSlides(images, pptxOutput, (slideErr, output) => {
            done(slideErr, output)
            rmdir(outputDir, (e) => e && errorLogger('Could not delete working directory:', e))
          })
        }
      })
    }, (err) => done(err) )
  }

  createSlides (imageFiles, pptFile, done) {
    pdfLogger('Adding images to slides')
    var pptx = officegen('pptx')
    imageFiles.forEach(i => {
      var slide = pptx.makeNewSlide();
      slide.addImage(i, {x: 0, y: 0, cx: "100%", cy: "100%"})
    })
    var out = fs.createWriteStream(pptFile)
    out.on('close', () => {
      pdfLogger('Created the PPTX file:', pptFile)
      done(null, pptFile)
    })
    pptx.generate(out)
  }

  _getStagingDirectory (stagingDir) {
    //return Promise.try( () => {
    return new Promise( (resolve,reject) => {
      if( stagingDir ){
        fs.stat(stagingDir, (err,s) => {
          if( err || !s.exists() ){
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
    return new Promise( (resolve,reject) => {
      fs.mkdtemp(`${os.tmpdir()}/pdf_ppt_`, (err,folder) => {
        if( err ) reject(err)
        pdfLogger('Created staging directory:', folder)
        resolve(folder)
      })
    })
  }

}

module.exports = Powerpoint
