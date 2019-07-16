const _ = require('lodash')
const Promise = require('bluebird')
const EventEmitter2 = require('eventemitter2').EventEmitter2

const fs = require('fs')
const execFile = require('child_process').execFile
const spawn = require('child_process').spawn
const os = require('os')
const path = require('path')

const { info } = require('./logger')
const { nowInMillis, elapsed, sortPages } = require('./util')

/** allows the caller to provide a path to the Inkscape Executable */
const INKSCAPE_PATH = process.env.PDF_PPT_INKSCAPE_PATH
/** allows the caller to provide a path to the MuPdf Executable */
const MUPDF_PATH = process.env.PDF_PPT_MUPDF_PATH
/** allows the caller to provide a path to the GhostScript Executable */
const GS_PATH = process.env.PDF_PPT_GSPATH
/** allows the caller to provide a path to the ImageMagick Executable */
const IM_PATH = process.env.PDF_PPT_IMPATH

let imageMagickConvert
let gsExecutable
let mupdfExecutable
let inkScapeExecutable

/** Going above 300 has a significant impact on performance
 * without much noticeable quality improvement */
const DPI_DENSITY_MAX = 300
/** 72 is the default (because of long-time standard) */
const DPI_DENSITY_DEFAULT = 72

const isWindowsPlatform = p => p === 'win32'
/**
 * Responsible for turning a PDF into a series of images.
 */
class Engine extends EventEmitter2 {
  /**
   * @param {String} [options.engine=ghostscript] options: mupdf, inkscape, ghostscript
   * @param {Boolean} [options.cropLastImage=false]
   * @param {Object} [options.convertOptions] ImageMagick conversion options
   * @param {Number} [options.convertOptions.density] must be < 300
   */
  constructor (options = {}) {
    super({
      wildcard: true, // Allow clients to listen based on wildcard expressions
      maxListeners: 10 // Node Default, set to 0 for Infinity
    })
    this._engine = options.engine
    this._cropLastImage = options.cropLastImage
    this._convertOptions = options.convertOptions || {}
    this._imgDir = options.imgDir
    this._pdfDir = options.pdfDir

    this._setExecutables()
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
    mupdfExecutable = MUPDF_PATH || 'mutool'
    inkScapeExecutable = INKSCAPE_PATH || 'inkscape'
    imageMagickConvert = IM_PATH || 'convert'
    // Windows
    if (isWindowsPlatform(p)) {
      gsExecutable = GS_PATH || 'gswin32c.exe'
      mupdfExecutable = MUPDF_PATH || 'mutool.exe'
      inkScapeExecutable = INKSCAPE_PATH || 'inkscape'
      imageMagickConvert = IM_PATH || 'convert.exe'
    }
  }

  /**
   * Start the conversion from PDF into an array of images.
   * @return {Promise}
   */
  convert (outputDir, files) {
    let convertPromise
    if (this._engine === 'inkscape') {
      convertPromise = this._convertWithInkscape(files)
    } else if (this._engine === 'mupdf') {
      convertPromise = this._convertWithMuPDF(files)
    } else {
      convertPromise = this._convertWithGhostScript(files)
    }
    return convertPromise
  }

  /**
   * GhostScript can be invoked directly, since ImageMagick just delegates to it
   *
   * @param pdfFiles
   * @param options
   * @private
   */
  _convertWithGhostScript (pdfFiles) {
    const start = nowInMillis()
    const co = this._getConvertOptions()

    const gsOpts = ['-q', '-dQUIET', '-dSAFER', '-sDEVICE=pngalpha', '-dMaxBitmap=500000000', `-r${co.density}`]

    // Get the image files for each PDF
    const gsErr = []

    const tasks = pdfFiles.map((pdfPath, pdfIndex) => {
      return new Promise((resolve) => {
        const imgPrefix = `img-${pdfIndex}-`
        const args = [...gsOpts, `-o "${imgPrefix}%d.png`, pdfPath]
        const gsStart = nowInMillis()
        execFile(gsExecutable, args, { cwd: this._imgDir }, (err, stdout, stderr) => {
          this.emit('done.gs.convert', { output: `${gsExecutable} ${args.join(' ')}`, time: elapsed(gsStart), error: err })
          if (err) {
            gsErr.push(err)
          }
          resolve()
        })
      })
    })

    return this._processImgConversionTasks(tasks, gsErr, start)
  }

  _convertWithInkscape (pdfFiles) {
    // Split the PDFs
    const splitTasks = this._getSplitTasks(pdfFiles)
    return Promise.all(splitTasks).then(() => {
      return this._readPdfDirectory().then(singlePagePdfFiles => {
        const sortedSinglePDFs = sortPages(singlePagePdfFiles)
        return this._executeInkscape(sortedSinglePDFs)
      })
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
      fs.readdir(this._pdfDir, (err, files) => {
        if (err) {
          reject(err)
        }
        resolve(files.map(f => `${this._pdfDir}/${f}`))
      })
    })
  }

  _getSplitTasks (pdfFiles) {
    return pdfFiles.map((pdfPath, pdfIndex) => {
      return new Promise((resolve, reject) => {
        const splitStart = nowInMillis()
        const args = [pdfPath, `${this._pdfDir}/pdf-${pdfIndex}-%d.pdf`]
        execFile('pdfseparate', args, (err, stdout, stderr) => {
          this.emit('done.pdf.separate', { output: `pdfseparate ${args.join(' ')}`, time: elapsed(splitStart), error: err })
          if (err) {
            reject(err)
          }
          resolve()
        })
      })
    })
  }

  _executeInkscape (sortedSinglePDFs) {
    const sortedImages = []

    const co = this._getConvertOptions()
    const commands = _.map(sortedSinglePDFs, (pdfFile) => {
      const pngFile = path.join(this._imgDir, `${path.basename(pdfFile, '.pdf')}.png`)
      sortedImages.push(pngFile)
      return [`-d ${co.density}`, `--export-png=${pngFile}`, pdfFile]
    })

    const inkTasks = this._getInkscapeExportTasks(commands)
    const start = nowInMillis()
    return Promise.all(inkTasks).then(() => {
      this.emit('done.inkscape.export.all', { time: elapsed(start) })
      return Promise.resolve(sortedImages)
    })
    // return this._spawnInkscapeShell(sortedImages, commands).then(() => sortedImages)
  }

  /**
   * @param {Array<String[]>>} commands List of Inkscape args per invocation, to be passed to execFile
   */
  _getInkscapeExportTasks (commands) {
    return commands.map(args => {
      return new Promise((resolve, reject) => {
        // Including timings here isn't useful because these promises run concurrently
        execFile(inkScapeExecutable, args, (err, stdout, stderr) => {
          this.emit('done.inkscape.export', { output: `${inkScapeExecutable} ${args.join(' ')}`, error: err })
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
            { output: msg, time: elapsed(inkCmdStart) })
          inkCmdStart = nowInMillis()
        }
      })
      inkProc.on('error', e => {
        info('Inkscape conversion failed:', e)
        reject(e)
      })
      inkProc.on('exit', () => {
        this.emit('done.inkscape.export.all', { time: elapsed(inkExecStart) })
        resolve()
      })
      // Run all the conversions in the shell
      const inkShellCmd = _.join(commands, ' \n') + ' \nquit\n'
      info('Inkscape shell commands:', inkShellCmd)
      const inkExecStart = nowInMillis()
      let inkCmdStart = inkExecStart
      inkProc.stdin.write(inkShellCmd)
    })
  }

  /**
   * Converts an array of input files (PNG or PDF) into an array of image files
   * that will be placed on powerpoint slides.
   *
   * @param files An array of PDF or PNG files
   * @returns {Promise<String[]>} A list of all the PNG files needed for slides
   */
  _convertWithMuPDF (files) {
    const start = nowInMillis()
    const co = this._getConvertOptions()

    // Get the image files for each PDF
    const errors = []

    const tasks = files.map((filePath, fileIndex) => {
      return new Promise((resolve) => {
        const imgPrefix = `img-${fileIndex}-`
        if (filePath.endsWith('.png')) { // Already have a PNG, just copy it
          fs.copyFile(filePath, `${this._imgDir}/${imgPrefix}0.png`, (err) => {
            if (err) {
              errors.push(err)
            }
            resolve()
          })
        } else { // Need to convert the PDF to PNG(s)
          const args = ['draw', `-r ${co.density}`, `-o ${imgPrefix}%d.png`, filePath]
          const muStart = nowInMillis()
          execFile(mupdfExecutable, args, { cwd: this._imgDir }, (err, stdout, stderr) => {
            this.emit('done.mupdf.convert', { output: `${mupdfExecutable} ${args.join(' ')}`, time: elapsed(muStart), error: err })
            if (err) {
              errors.push(err)
            }
            resolve()
          })
        }
      })
    })
    return this._processImgConversionTasks(tasks, errors, start)
  }

  /**
   * @param {Array} tasks list of promises
   * @param {Array} errors
   * @param {Number} startedAt timestamp of when conversion started
   * @returns {Promise.<String[]>} sorted images ready for pptx slides
   */
  _processImgConversionTasks (tasks, errors, startedAt) {
    return Promise.all(tasks).then(() => {
      if (!_.isEmpty(errors)) {
        this.emit('err.png.all', { error: errors, time: elapsed(startedAt) })
        return Promise.reject(errors)
      } else {
        const imagesFiles = fs.readdirSync(this._imgDir).map(f => `${this._imgDir}/${f}`)
        const sortedImages = sortPages(imagesFiles)
        this.emit('done.png.all', { output: sortedImages, time: elapsed(startedAt) })
        return this._cropLastImages(sortedImages).then(() => sortedImages)
      }
    })
  }

  _getConvertOptions () {
    const co = this._convertOptions
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
    if (!this._cropLastImage) {
      return Promise.resolve()
    }

    const start = nowInMillis()
    const convertErrors = []

    const conversions = this._getLastImageFiles(sortedImages).map((img) => {
      return new Promise((resolve) => {
        /*
         * -gravity south is needed for the chop instruction so we get the bottom edge
         *
         * border options are used to add a single pixel of transparency before trim autocrops the image.
         * Without this the top is trimmed and we don't want that
         *
         * -trim leaves the partially transparent pixel at the bottom so we have
         * to chop it off, but we know there is only one pixel after the trim.
         */
        const args = [`-gravity South`, `-bordercolor none`, `-border 1`, `-trim`, `+repage`, `-chop 0x1`, img]
        var imStart = nowInMillis()
        execFile(imageMagickConvert, args, (err, stdout, stderr) => {
          this.emit('done.im.convert', { output: `${imageMagickConvert} ${args.join(' ')}`, time: elapsed(imStart) })
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
            this.emit('err.im.convert', { error: convertErrors, time: elapsed(start) })
          } else {
            this.emit('done.im.convert.all', { time: elapsed(start) })
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
}

module.exports = Engine
