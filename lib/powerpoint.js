'use strict'

var EventEmitter2 = require('eventemitter2').EventEmitter2
var exec = require('child_process').exec
var spawn = require('child_process').spawn
var fs = require('fs')
var os = require('os')
var path = require('path')

const _ = require('lodash')
var officegen = require('officegen')
var Promise = require('bluebird')
var rmdir = require('rimraf')

// Loggers
var debug = require('debug')
var pdfLogger = debug('pdfppt:app')
var debugLogger = debug('pdfppt:debug')
pdfLogger.log = console.log.bind(console)
debugLogger.log = console.log.bind(console)

/** allows the caller to provide a path to the Inkscape Executable */
const INKSCAPE_PATH = process.env.PDF_PPT_INKSCAPE_PATH
/** allows the caller to provide a path to the MuPdf Executable */
const MUPDF_PATH = process.env.PDF_PPT_MUPDF_PATH
/** allows the caller to provide a path to the GhostScript Executable */
const GS_PATH = process.env.PDF_PPT_GSPATH
/** allows the caller to provide a path to the ImageMagick Executable */
const IM_PATH = process.env.PDF_PPT_IMPATH

let gsExecutable
let mupdfExecutable
let inkScapeExecutable
let imageMagickConvert

/** Going above 300 has a significant impact on performance
 * without much noticeable quality improvement */
const DPI_DENSITY_MAX = 300
/** 72 is the default (because of long-time standard) */
const DPI_DENSITY_DEFAULT = 72

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
    this._setExecutables()
    this.options = _.merge({}, defaultOptions, options)
    process.nextTick(() => {
      this.emit('options', { output: JSON.stringify(this.options) })
    })
  }

  /**
   * Inspects any env property that may have been set, and changes the default
   * based on operating system if not set.
   * @private
   */
  _setExecutables () {
    const p = os.platform()
    // Linux
    gsExecutable = GS_PATH || 'gs'
    mupdfExecutable = MUPDF_PATH || 'mudraw'
    inkScapeExecutable = INKSCAPE_PATH || 'inkscape'
    imageMagickConvert = IM_PATH || 'convert'
    // Windows
    if (p === 'win32') {
      gsExecutable = GS_PATH || 'gswin32c.exe'
      mupdfExecutable = MUPDF_PATH || 'mutool.exe draw'
      imageMagickConvert = IM_PATH || 'convert.exe'
      inkScapeExecutable = INKSCAPE_PATH || 'inkscape'
    }
    // Mac OS
    if (p === 'darwin') {
      mupdfExecutable = MUPDF_PATH || 'mutool draw'
    }
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

    // pdfLogger('options:', opts)

    this._getStagingDirectory(opts.stagingDir)
      .then((outputDir) => {
        this.imgDir = path.resolve(outputDir, 'img')
        this.pdfDir = path.resolve(outputDir, 'pdf')
        let convertPromise
        if (this.options.engine === 'inkscape') {
          convertPromise = this._convertWithInkscape(outputDir, pdfFiles, options)
        } else if (this.options.engine === 'mupdf') {
          convertPromise = this._convertWithMuPDF(outputDir, pdfFiles, options)
        } else {
          convertPromise = this._convertWithGhostScript(outputDir, pdfFiles, options)
        }
        convertPromise.then(sortedImages => {
          this._createPowerpoint(outputDir, sortedImages, callback)
        })
      }, (err) => callback(err))
      .catch(err => callback(err))
  }

  _convertWithInkscape (outputDir, pdfFiles, options) {
    // Split the PDFs
    const splitTasks = this._getSplitTasks(pdfFiles)
    return Promise.all(splitTasks).then(() => {
      return this._readPdfDirectory().then(singlePagePdfFiles => {
        const sortedSinglePDFs = this._sortPages(singlePagePdfFiles)
        return this._executeInkscape(sortedSinglePDFs, options)
      })
    })
  }

  _executeInkscape (sortedSinglePDFs, options) {
    const sortedImages = []

    const co = this._getConvertOptions(options)
    const commands = _.map(sortedSinglePDFs, (pdfFile) => {
      const pngFile = `${path.basename(pdfFile, '.pdf')}.png`
      const pngPath = path.join(this.imgDir, pngFile)
      sortedImages.push(pngPath)
      return `-d ${co.density} --export-png=${pngPath} ${pdfFile}`
    })

    const inkTasks = this._getInkscapeExportTasks(commands)
    const start = this.nowInMillis()
    return Promise.all(inkTasks).then(() => {
      this.emit('done.inkscape.export.all', {time: this.elapsed(start)})
      return Promise.resolve(sortedImages)
    })
    // return this._spawnInkscapeShell(sortedImages, commands).then(() => sortedImages)
  }

  _getInkscapeExportTasks (commands) {
    return commands.map((cmd) => {
      return new Promise((resolve, reject) => {
        const fullCmd = `${inkScapeExecutable} ${cmd}`
        // Including timings here isn't useful because these promises run concurrently
        exec(fullCmd, (err, stdout, stderr) => {
          this.emit('done.inkscape.export', { output: fullCmd, error: err })
          if (err) {
            reject(err)
          }
          resolve()
        })
      })
    })
  }

  /**
   * Executes inkscape export commands in a single reusable shell.
   * In theory this should be more efficient, but it runs serially and ends
   * up taking a lot longer.
   *
   * @param {Array} commands inkscape export commands
   * @returns {Promise}
   * @private
   */
  _spawnInkscapeShell (commands) {
    return new Promise((resolve, reject) => {
      const inkProc = spawn(inkScapeExecutable, ['--shell'])
      inkProc.stdout.on('data', d => {
        // Each export shell command writes 3 lines to stdout, this is all we have to
        // signal that a single conversion was completed.  The `Bitmap saved` line is the
        // most contextual (and last) line, so that is what is logged
        const msg = d.toString()
        if (_.startsWith(msg, 'Bitmap saved as:')) {
          this.emit('done.inkscape.export',
            {output: msg, time: this.elapsed(inkCmdStart)})
          inkCmdStart = this.nowInMillis()
        }
      })
      inkProc.on('error', e => {
        pdfLogger('Inkscape conversion failed:', e)
        reject(e)
      })
      inkProc.on('exit', () => {
        this.emit('done.inkscape.export.all', {time: this.elapsed(inkExecStart)})
        resolve()
      })
      // Run all the conversions in the shell
      const inkShellCmd = _.join(commands, ' \n') + ' \nquit\n'
      pdfLogger('Inkscape shell commands:', inkShellCmd)
      let inkExecStart = this.nowInMillis()
      let inkCmdStart = inkExecStart
      inkProc.stdin.write(inkShellCmd)
    })
  }

  /**
   * Reads the contents of the staging directory that contains the PDF
   * files after they are split using pdfseparate.  This is because we have
   * no idea how many files there will be, and they need to be sorted properly.
   *
   * @returns {Promise} fulfilled with list of filenames
   */
  _readPdfDirectory () {
    return new Promise((resolve, reject) => {
      fs.readdir(this.pdfDir, (err, files) => {
        if (err) {
          reject(err)
        }
        resolve(files.map(f => `${this.pdfDir}/${f}`))
      })
    })
  }

  _getSplitTasks (pdfFiles) {
    return pdfFiles.map((pdfPath, pdfIndex) => {
      return new Promise((resolve, reject) => {
        const splitCmd = `pdfseparate ${pdfPath} ${this.pdfDir}/pdf-${pdfIndex}-%d.pdf`
        const splitStart = this.nowInMillis()
        exec(splitCmd, (err, stdout, stderr) => {
          this.emit('done.pdf.separate', { output: splitCmd, time: this.elapsed(splitStart), error: err })
          if (err) {
            reject(err)
          }
          resolve()
        })
      })
    })
  }

  /**
   * GhostScript can be invoked directly, since ImageMagick just delegates to it
   *
   * @param outputDir
   * @param pdfFiles
   * @param options
   * @private
   */
  _convertWithGhostScript (outputDir, pdfFiles, options) {
    const start = this.nowInMillis()
    const co = this._getConvertOptions(options)

    const gsCmdRoot = `"${gsExecutable}" -q -dQUIET -dSAFER -sDEVICE=pngalpha -dMaxBitmap=500000000 -r${co.density}`

    // Get the image files for each PDF
    let gsErr = []

    let tasks = pdfFiles.map((pdfPath, pdfIndex) => {
      return new Promise((resolve) => {
        const imgPrefix = `img-${pdfIndex}-`
        const gsCmd = gsCmdRoot + ` -o "${this.imgDir}/${imgPrefix}%d.png" "${pdfPath}"`
        const gsStart = this.nowInMillis()
        exec(gsCmd, (err, stdout, stderr) => {
          this.emit('done.gs.convert', { output: gsCmd, time: this.elapsed(gsStart), error: err })
          if (err) {
            gsErr.push(err)
          }
          resolve()
        })
      })
    })

    return this.processImgConversionTasks(tasks, gsErr, start)
  }

  /**
   * @param {Array} tasks list of promises
   * @param {Array} errors
   * @param {Number} startedAt timestamp of when conversion started
   * @returns {Promise.<String[]>} sorted images ready for pptx slides
   */
  processImgConversionTasks (tasks, errors, startedAt) {
    return Promise.all(tasks).then(() => {
      if (!_.isEmpty(errors)) {
        this.emit('err.png.all', {error: errors, time: this.elapsed(startedAt)})
        return Promise.reject(errors)
      } else {
        const imagesFiles = fs.readdirSync(this.imgDir).map(f => `${this.imgDir}/${f}`)
        const sortedImages = this._sortPages(imagesFiles)
        this.emit('done.png.all', {output: sortedImages, time: this.elapsed(startedAt)})
        return this._cropLastImages(sortedImages).then(() => sortedImages)
      }
    })
  }

  _convertWithMuPDF (outputDir, pdfFiles, options) {
    const start = this.nowInMillis()
    const co = this._getConvertOptions(options)

    const cmdRoot = `${mupdfExecutable} -r ${co.density}`

    // Get the image files for each PDF
    let errors = []

    let tasks = pdfFiles.map((pdfPath, pdfIndex) => {
      return new Promise((resolve) => {
        const imgPrefix = `img-${pdfIndex}-`
        const cmd = cmdRoot + ` -o ${this.imgDir}/${imgPrefix}%d.png ${pdfPath}`
        const muStart = this.nowInMillis()
        exec(cmd, (err, stdout, stderr) => {
          this.emit('done.mupdf.convert', { output: cmd, time: this.elapsed(muStart), error: err })
          if (err) {
            errors.push(err)
          }
          resolve()
        })
      })
    })
    return this.processImgConversionTasks(tasks, errors, start)
  }

  _createPowerpoint (outputDir, sortedImages, callback) {
    var pptxOutput = path.resolve(outputDir, `output_${process.hrtime()[1]}.pptx`)
    this._aggregateSlides(sortedImages, pptxOutput, this.imgDir, callback)
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
   * Acts on image files in place.
   *
   * If the last page of a PDF did not represent a full page in the DOM it is
   * possible for there to be a boundary that has a transparency level that
   * results in a horizontal line being visible when the image is added to the
   * powerpoint slide.
   *
   * This function crops this transparency and removes the line, by invoking
   * ImageMagick
   *
   * @param sortedImages
   *
   * @returns {Promise} when all conversions have been done, and events emitted
   * Any errors will be logged, but a rejection will not occur.  It is better
   *   to get an output with a line than no output at all.
   *
   * @private
   */
  _cropLastImages (sortedImages) {
    if (!this.options.cropLastImage) {
      return Promise.resolve()
    }

    const start = this.nowInMillis()
    let convertErrors = []

    const conversions = this._getLastImageFiles(sortedImages).map((img) => {
      return new Promise((resolve) => {
        /*
         * -gravity south is needed for the chop instruction so we get the bottom edge
         * border options are used to add a single pixel of transparency before trim autocrops the image. Without this the top is trimmed and we don't want that
         * -trim leaves the partially transparent pixel at the bottom so we have to chop it off, but we know there is only one pixel after the trim.
         */
        const command = `${imageMagickConvert} ${img} -gravity South -bordercolor none -border 1 -trim +repage -chop 0x1 ${img}`
        var imStart = this.nowInMillis()
        exec(command, (err, stdout, stderr) => {
          this.emit('done.im.convert', { output: command, time: this.elapsed(imStart) })
          if (err) {
            convertErrors.push(err)
          }
          if (stderr) {
            convertErrors.push(stderr)
          }
          resolve()
        })
      })
    })

    return Promise.all(conversions)
                  .then(() => {
                    return new Promise((resolve) => {
                      if (!_.isEmpty(convertErrors)) {
                        this.emit('err.im.convert', {error: convertErrors, time: this.elapsed(start)})
                      } else {
                        this.emit('done.im.convert.all', {time: this.elapsed(start)})
                      }
                      resolve()
                    })
                  })
  }

  /**
   * @param sortedImages sorted by file and page/image number
   * @returns {Array} of files from sortedImages that are the last image for
   *   each file
   * @private
   */
  _getLastImageFiles (sortedImages) {
    let lastFile
    return sortedImages.reduce((acc, val, i, arr) => {
      const fileAndPage = /.*[pdf|img]-(\d*)-(\d*).*/.exec(val)
      const file = fileAndPage[1]

      const nextFile = lastFile !== undefined && file !== lastFile
      const lastImage = i === arr.length - 1

      const lastImages = [...acc]
      if (nextFile) {
        lastImages.push(arr[i - 1])
      }
      if (lastImage) {
        lastImages.push(val)
      }
      lastFile = file
      return lastImages
    }, [])
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
        var start = this.nowInMillis()
        rmdir(imgDir, (err) => {
          if (err) {
            this.emit('done.png.clean', {output: imgDir, time: this.elapsed(start), error: err})
            pdfLogger(this.options.jobId, 'Could not delete working directory:', imgDir, err)
          }
        })
      }
    })
  }

  _createSlides (imageFiles, pptFile, done) {
    var start = this.nowInMillis()
    var pptx = officegen('pptx')
    var d = this.options.dimensions
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

  _sortPages (imageFiles) {
    // Example: /var/folders/dr/f1q4znd96xv8wp82y4cfgg700000gn/T/pdf_ppt_5tz0dw/img/img-5-10.png
    // File = 5, Page = 10
    const rex = /.*(img|pdf)-(\d*)-(\d*).*/
    return imageFiles.sort((a, b) => {
      let aGrps = rex.exec(a)
      let bGrps = rex.exec(b)
      // PDF File Sequence + Page Sequence Comparison
      const fileGrp = 2
      const pageGrp = 3
      let fileComp = aGrps[fileGrp] - bGrps[fileGrp]
      if (fileComp === 0) {
        return aGrps[pageGrp] - bGrps[pageGrp]
      }
      return fileComp
    })
  }

  _savePresentationFile (pptFile, done, pptx) {
    var start = this.nowInMillis()
    var out = fs.createWriteStream(pptFile)
    out.on('close', () => {
      this.emit('done.pptx.saved', {output: pptFile, time: this.elapsed(start)})
      done(null, pptFile)
    })
    pptx.generate(out)
  }

  _getStagingDirectory (stagingDir) {
    return new Promise((resolve, reject) => {
      if (stagingDir) {
        fs.stat(stagingDir, (err, s) => {
          if (err || !s.isDirectory()) {
            pdfLogger(this.options.jobId, 'staging directory:', stagingDir, 'does not exist, creating a new one')
            return this._createTempStagingDirectory()
          } else {
            this._createImageDirectory(stagingDir, reject, resolve)
          }
        })
      } else {
        resolve(this._createTempStagingDirectory())
      }
    })
  }

  _createTempStagingDirectory () {
    return new Promise((resolve, reject) => {
      fs.mkdtemp(path.join(os.tmpdir(), 'pdf_ppt_'), (err, folder) => {
        if (err) reject(err)
        this._createImageDirectory(folder, reject, resolve)
      })
    })
  }

  _createImageDirectory (folder, reject, resolve) {
    fs.mkdir(path.resolve(folder, 'img'), (err) => {
      if (err) reject(err)
      fs.mkdir(path.resolve(folder, 'pdf'), (err) => {
        if (err) reject(err)
        resolve(folder)
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
