'use strict'

var fs = require('fs')

var officegen = require('officegen')
var PDF2Images = require('pdf2images-multiple')

class Powerpoint {

  // Pass through options:
  /*
   - convertOptions
   - slideImageOptions
   -
   */

  convertPDFToImages (pdfFile, imgOutput, done) {
    var convert_options = {
      '-density' : 600, //DPI
      //'-resize' : '800x600',
      '-quality' : 100
    }
    var convert_operators = {
      //'-trim': '',
      //'-sharpen' : '0x1.0'
    }

    var pdf2images = PDF2Images(pdfFile, {
      convert_options: convert_options, //optional
      convert_operators: convert_operators, //optional
      //output_dir: './media/', //optional
      ext: 'png', //optional, png is the default value
      gm: false //Use GraphicksMagic //optional, false is the default value
    })

    console.log('converting', pdfFile, 'into images...')

    pdf2images.pdf.convert( (err, image_path) => {
      // Will it be faster to add images here?
    }, (err, images) => {
      if( err ){
        console.log(err)
      } else {
        console.log('creating slides')
        this.createSlides(images, 'slides.pptx', () => done())
      }
    })
  }

  createSlides(imageFiles, pptFile, done){
    console.log('Adding images to slides')
    var pptx = officegen ( 'pptx' )
    imageFiles.forEach( i => {
      var slide = pptx.makeNewSlide ();
      slide.addImage(i, {x:0, y:0, cx: "100%", cy: "100%"})
    })
    var out = fs.createWriteStream(pptFile)
    out.on ( 'close', () => {
      console.log ( 'Created the PPTX file!' )
      done()
    })
    pptx.generate (out)
  }

}

module.exports = Powerpoint
