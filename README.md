[![Build Status](https://travis-ci.org/SpiderStrategies/pdf-powerpoint.svg?branch=master)](https://travis-ci.org/SpiderStrategies/pdf-powerpoint)
[![npm version](https://badge.fury.io/js/pdf-powerpoint.svg)](https://badge.fury.io/js/pdf-powerpoint)

[![NPM](https://nodei.co/npm/pdf-powerpoint.png?downloads=true&stars=true)](https://nodei.co/npm/pdf-powerpoint/)

## PDF to Powerpoint Converter

A NPM module that accepts one or more PDF files and converts them into Powerpoint slides.

- GhostScript is used to transform each page of a PDF into a PNG image.
- Each single images is added to a slide in the powerpoint presentation.
- Slides are in the order of the PDFs passed in the array

**Supported Runtimes:**  Node > 5.10.0

**Required packages:** 
- Debian: `apt-get install -y ghostscript`
- OSX: `brew install ghostscript`

### Usage

```javascript
import {Powerpoint} from 'pdf-powerpoint'
const p = new Powerpoint()

p.convertPDFToPowerpoint('input.pdf', [options,] (err, result) => {
  //Do something with the result (filepath to output) 
})
  
```

### Options

* `stagingDir` - This is where the pptx file will be generated.  
  - Images go in `stagingDir/img` and are automatically removed once the powerpoint file is generated.
  - If not provided the default is to use `fs.mkdtemp('${os.tmpdir()}/pdf_ppt_')` to generate a random temp directory
* `convertOptions` - These are used for Ghostscript conversion
  - `density` - specifies the PPI setting for the output image
    - default: 72, maximum value allowed is 300
  
### Events

Events are emitted for any client wishing to capture timings or react to incremental artifacts.

The [EventEmitter2](https://www.npmjs.com/package/eventemitter2) library is used, which means you
 can write a single listener for all events if you wish.
 
Events emit an object that may have the following properties:
* `error` - if an error occurred
* `time` - if the event marks the end of a corresponding start event
* `output` - If there is a PNG or PPTX file generated from the event

#### Event Names

1. `done.gs.convert`- `output` is the GhostScript command that was executed
1. `err.png.all` 
1. `done.png.all` - `output` is an array of paths to images generated from PDF
1. `done.png.clean` - `output` is the image directory that was deleted
1. `done.pptx.creation` - powerpoint is complete in memory, all images added to slides
1. `done.pptx.saved` - `output` is the pptFile

### Logging

Debug is used for logging and there are three namespaces you can enable.

* pdfppt:app
* pdfppt:debug

This can be turned on by setting `DEBUG=pdfppt:*`, read more about [Debug here](https://www.npmjs.com/package/debug)

### Implementation
 
#### ImageMagick delegates to GhostScript for PDF -> PNG conversion

- ImageMagick: `convert -density 72 -quality 100 -verbose  '/var/folders/dr/f1q4znd96xv8wp82y4cfgg700000gn/T/833198680xmyTzU/output.pdf[4]' '/var/folders/dr/f1q4znd96xv8wp82y4cfgg700000gn/T/pdf_ppt_Tl9eSm/img/output-4.png'`
- GhostScript: `'gs' -q -dQUIET -dSAFER -dBATCH -dNOPAUSE -dNOPROMPT -dMaxBitmap=500000000 -dAlignToPixels=0 -dGridFitTT=2 '-sDEVICE=pngalpha' -dTextAlphaBits=4 -dGraphicsAlphaBits=4 '-r72x72' -dFirstPage=5 -dLastPage=5 '-sOutputFile=/var/tmp/magick-94224ozuZS3iFphAj%d' '-f/var/tmp/magick-94224zWXBFMw8ZiEA' '-f/var/tmp/magick-9422413LS3T1dhoL4'`

#### So GhostScript is used directly 

*Note:* You must ensure that GhostScript is installed on your system, it is not included with this package.
If you have GhostScript installed globally on your system it should be located automatically, but if not you can provide the path to the GhostScript executable by setting the `PDF_PPT_GSPATH` environment variable.

[GhostScript Option Documentation](https://ghostscript.com/doc/current/Use.htm)

The following command is generated: `gs -q -dQUIET -sDEVICE=pngalpha -r150 -o outputFile-%d.png`

- As a convenient shorthand you can use the `-o option` followed by the output file specification as discussed above. The -o option also sets the `-dBATCH` and `-dNOPAUSE` options.
- `-q` Quiet startup: suppress normal startup messages, and also do the equivalent of -dQUIET.
- `-dQUIET` Suppresses routine information comments on standard output.
- `-sDEVICE=pngalpha` 
- `-r[XResxYRes]` Useful for controlling the density of pixels when rasterizing to an image file. It is the requested number of dots (or pixels) per inch. Where the two resolutions are same, as is the common case, you can simply use -rres.
