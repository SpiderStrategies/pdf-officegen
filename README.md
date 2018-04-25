[![Build Status](https://travis-ci.org/SpiderStrategies/pdf-powerpoint.svg?branch=master)](https://travis-ci.org/SpiderStrategies/pdf-powerpoint)
[![npm version](https://badge.fury.io/js/pdf-powerpoint.svg)](https://badge.fury.io/js/pdf-powerpoint)

[![NPM](https://nodei.co/npm/pdf-powerpoint.png?downloads=true&stars=true)](https://nodei.co/npm/pdf-powerpoint/)

# PDF to Powerpoint Converter

A NPM module that accepts one or more PDF files and converts them into one of the following:
1. Powerpoint slides (.pptx)
2. Office Word Document (.docx)

### General workflow
- A rendering engine is used to transform each page of a PDF into a PNG image.
- Each single images is added to a slide in the powerpoint presentation.
- Slides are in the order of the PDFs passed in the array

### PDF Rendering engines
Based on the requirements of your application, one rendering engine may be more appropriate
than another.  This library currently supports three options.  In all cases, you must ensure
the binaries are installed for your runtime, they are not packaged with this module.

It is recommended you weigh the runtime performance and output quality of each engine for
the content you are converting.

1. GhostScript - Converts a PDF into PNGs with one command per PDF
    - Debian: `apt-get install -y ghostscript`
    - OSX: `brew install ghostscript`
1. MuPDF - Converts a PDF into PNGs with one command per PDF
    - Debian: `apt-get install -y mupdf-tools`
    - OSX: `brew install ghostscript`
1. Inkscape - Separates PDFs into single page PDFs and then converts each PDF into PNG
    - Debian: `apt-get install -y inkscape`
    - OSX: `brew install inkscape`

**Supported Runtimes:**  Node > 5.10.0

## Usage

```javascript
import {Powerpoint, Word} from 'pdf-powerpoint'
const p = new Powerpoint([options])
````

#### Constructor Options
* `clean` - set to false if intermediate image files should be left on the filesystem.
* `jobId` - if provided, this will be included in any logging output
* `cropLastImage` requires ImageMagick `convert` to be on the path.  Will crop the last pdf image before placing on slide, sometimes a line would show up if the last PDF page was a partial page.
* `dimensions`
    - width - of slides in pixels (default: 800)
    - height - of slides in pixels (default: 600)
    - type - options: '35mm' 'A3' 'A4', 'B4ISO' 'B4JIS' 'B5ISO' 'B5JIS' 'banner' 'custom' 'hagakiCard' 'ledger', 'letter' 'overhead' 'screen16x10' 'screen16x9' 'screen4x3' (default)


```javascript
p.convertFromPdf('input.pdf', [options,] (err, result) => {
  //Do something with the result (filepath to output) 
})
```

#### Convert Options

- `engine`
  - 'ghostscript' (default)
  - 'mupdf'
  - 'inkscape'
- `stagingDir` - This is where the pptx file will be generated.  
  - Images go in `stagingDir/img` and are automatically removed once the powerpoint file is generated.
  - If not provided the default is to use `fs.mkdtemp('${os.tmpdir()}/pdf_ppt_')` to generate a random temp directory
- `convertOptions` - These are used for Ghostscript conversion
  - `density` - specifies the PPI setting for the output image
    - default: 72, maximum value allowed is 300
  
### Events

Events are emitted for any client wishing to capture timings or react to incremental artifacts.

The [EventEmitter2](https://www.npmjs.com/package/eventemitter2) library is used, which means you
 can write a single listener for all events if you wish.
 
Events emit an object that may have the following properties:
- `error` - if an error occurred
- `time` - if the event marks the end of a corresponding start event
- `output` - If there is a PNG or PPTX file generated from the event

#### Event Names

- `err.png.all` 
- `done.png.all` - `output` is an array of paths to images generated from PDF
- `done.png.clean` - `output` is the image directory that was deleted
- `done.[pptx|docx].creation` - powerpoint is complete in memory, all images added to slides
- `done.[pptx|docx].saved` - `output` is the pptFile
- `err.[pptx|docx]` - `error` is the error the was thrown from officegen

##### Inkscape Engine
- `done.pdf.separate` - `output` is the command executed
- `done.inkscape.export` - after each inkscape conversion completes, `output` is the command that was executed
- `done.inkscape.export.all` - after all inkscape conversions are complete

##### GhostScript Engine
- `done.gs.convert`- `output` is the GhostScript command that was executed
- Only when `cropLastImage` option is set
    - `done.im.convert` - after the last image of each PDF is converted 
    - `done.im.convert.all` - after all images are cropped
    - `err.im.convert` - if any of the image cropping operations fails

##### MuPDF Engine
1. `done.mupdf.convert` - `output` is the MuPDF (mudraw) command that was executed

### Logging

Debug is used for logging and there are three namespaces you can enable.

* pdfppt:app
* pdfppt:debug

This can be turned on by setting `DEBUG=pdfppt:*`, read more about [Debug here](https://www.npmjs.com/package/debug)

## Developer Guide
 
This library originally used ImageMagick but it was discovered that ImageMagick delegates to GhostScript for PDF -> PNG conversion, so GhostScript is used directly

- ImageMagick: `convert -density 72 -quality 100 -verbose  '/tmp/output.pdf[4]' '/tmp/img/output-4.png'`
- Results in (GhostScript): `'gs' -q -dQUIET -dSAFER -dBATCH -dNOPAUSE -dNOPROMPT -dMaxBitmap=500000000 -dAlignToPixels=0 -dGridFitTT=2 '-sDEVICE=pngalpha' -dTextAlphaBits=4 -dGraphicsAlphaBits=4 '-r72x72' -dFirstPage=5 -dLastPage=5 '-sOutputFile=/tmp/magick-94224ozuZS3iFphAj%d' '-f/tmp/magick-94224zWXBFMw8ZiEA' '-f/tmp/magick-9422413LS3T1dhoL4'`

[GhostScript Option Documentation](https://ghostscript.com/doc/current/Use.htm)

This module uses the following command: `gs -q -dQUIET -sDEVICE=pngalpha -r150 -o outputFile-%d.png`

#### GhostScript Tips (not this module's API)
- As a convenient shorthand you can use the `-o option` followed by the output file specification as discussed above. The -o option also sets the `-dBATCH` and `-dNOPAUSE` options.
- `-q` Quiet startup: suppress normal startup messages, and also do the equivalent of -dQUIET.
- `-dQUIET` Suppresses routine information comments on standard output.
- `-sDEVICE=pngalpha` 
- `-r[XResxYRes]` Useful for controlling the density of pixels when rasterizing to an image file. It is the requested number of dots (or pixels) per inch. Where the two resolutions are same, as is the common case, you can simply use -rres.
