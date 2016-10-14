import {test} from 'ava'
import Powerpoint from '../lib/powerpoint'

const p = new Powerpoint()

test.cb('image is resized to fit on slide', t => {

  p.convertPDFToImages('google-l.pdf', 'pdf.png', (result) => {
    console.log(result)
    t.end()
  })

})