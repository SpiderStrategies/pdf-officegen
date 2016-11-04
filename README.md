[![Build Status](https://travis-ci.org/SpiderStrategies/pdf-powerpoint.svg?branch=master)](https://travis-ci.org/SpiderStrategies/pdf-powerpoint)
[![npm version](https://badge.fury.io/js/pdf-powerpoint.svg)](https://badge.fury.io/js/pdf-powerpoint)

[![NPM](https://nodei.co/npm/pdf-powerpoint.png?downloads=true&stars=true)](https://nodei.co/npm/pdf-powerpoint/)

## PDF to Powerpoint Converter

A NPM module that accepts one or more PDF files and converts them into Powerpoint slides.

- ImageMagick is used to transform each page of a PDF into a PNG image.
- Each single images is added to a slide in the powerpoint presentation.
- Slides are in the order of the PDFs passed in the array

**Supported Runtimes:**  Node > 5.10.0

**Required packages:** 
- Debian: `apt-get install -y imagemagick ghostscript poppler-utils GraphicsMagick`
- OSX: `brew install imagemagick poppler`

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
* `convertOptions` - These are used for ImageMagick conversion
  - `density` - specifies the PPI setting for the output image
    - default: 72, maximum value allowed is 300
  
### Events

Events are emitted for any client wishing to capture timings or react to incremental artifacts.

The [EventEmitter2](https://www.npmjs.com/package/eventemitter2) library is used, which means you
 can write a single listener for all events if you wish.
 
Events emit an object that may have the following properties:
* `error` - if an error occured
* `time` - if the event marks the end of a corresponding start event
* `output` - If there is a PNG or PPTX file generated from the event

#### Event Names

1. `err.png.single`
1. `done.png.single` - `output` is the path to the png file
1. `err.png.all` 
1. `done.png.all` - `output` is an array of paths to images generated from PDF
1. `done.pptx.creation` - powerpoint is complete in memory, all images added to slides
1. `done.pptx.saved` - `output` is the pptFile

### Logging

Debug is used for logging and there are three namespaces you can enable.

* pdfppt:error
* pdfppt:app
* pdfppt:debug

This can be turned on by setting `DEBUG=pdfppt:*`, read more about [Debug here](https://www.npmjs.com/package/debug)