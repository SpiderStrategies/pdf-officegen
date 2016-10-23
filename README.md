[![Build Status](https://travis-ci.org/SpiderStrategies/pdf-powerpoint.svg?branch=master)](https://travis-ci.org/SpiderStrategies/pdf-powerpoint)
[![npm version](https://badge.fury.io/js/pdf-powerpoint.svg)](https://badge.fury.io/js/pdf-powerpoint)

[![NPM](https://nodei.co/npm/pdf-powerpoint.png?downloads=true&stars=true)](https://nodei.co/npm/pdf-powerpoint/)

## PDF to Powerpoint Converter

A NPM module that accepts one or more PDF files and converts them into Powerpoint slides.

- ImageMagick is used to transform each page of a PDF into a PNG image.
- Each single images is added to a slide in the powerpoint presentation.
- Slides are in the order of the PDFs passed in the array

**Supported Runtimes:**  Node > 5.10.0


### Usage

```
import {Powerpoint} from 'pdf-powerpoint'
const p = new Powerpoint()


p.convertPDFToImages('input.pdf', 'output.pptx', (result) => {
  //Do something with the result (filepath to output) 
})
  
```