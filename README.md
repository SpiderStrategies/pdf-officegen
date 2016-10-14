## PDF to Powerpoint Converter

A NPM module that accepts one or more PDF files and converts them into Powerpoint slides.

- ImageMagick is used to transform each page of a PDF into a PNG image.
- Each single images is added to a slide in the powerpoint presentation.
- Slides are in the order of the PDFs passed in the array

### Usage

```
import {Powerpoint} from 'pdf-powerpoint'
const p = new Powerpoint()


p.convertPDFToImages('input.pdf', 'output.pptx', (result) => {
  //Do something 
})

  
```