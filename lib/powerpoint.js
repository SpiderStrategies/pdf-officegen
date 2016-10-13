var fs = require('fs')

var officegen = require('officegen')
var PDF2Images = require('pdf2images-multiple')

class Powerpoint {

  convertPDFToImages (pdfFile, imgOutput, done) {
    console.log('converting ', pdfFile)
    var convert_options = {
      '-trim': '',
      '-density' : 150,
      '-quality' : 100,
      '-sharpen' : '0x1.0'
    }

    var pdf2images = PDF2Images(pdfFile, {
      convert_options: convert_options, //optional
      // output_dir: './media/', //optional
      ext: 'png', //optional, png is the default value
      gm: false //Use GraphicksMagic //optional, false is the default value
    })

    pdf2images.pdf.convert((err, image_path) => {
      //Do something when convert every single page.
    }, (err, images) => {
      this.createSlides(images, 'slides.pptx', () => done())
    })
  }

  createSlides(imageFiles, pptFile, done){
    console.log('Adding images to slides')
    var pptx = officegen ( 'pptx' )
    imageFiles.forEach( i => {
      var slide = pptx.makeNewSlide ();
      slide.addImage(i)
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
