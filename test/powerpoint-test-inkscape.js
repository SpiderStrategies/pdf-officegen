import path from 'path'
import {test} from 'ava'
import Powerpoint from '../lib/powerpoint'

const p = new Powerpoint({
  engine: 'inkscape',
  clean: false
})

/**
 * Won't run with 'npm test'.  This is meant as test harness to be run manually
 * for validating inkscape end to end when making changes.  Developer must
 * provide the pdf below, it should not be included in SCM.
 */
test.cb('pdf is separated into images and placed on slides', t => {
  const options = {
    convertOptions: {
      density: 150
    }
  }
  p.onAny(console.log)
  p.convertPDFToPowerpoint([path.join(__dirname, 'putPdfHere.pdf')], options, (err, result) => {
    if (err) {
      t.fail('Did not complete successfully')
      console.log(err)
    }
    console.log(result)
    t.end()
  })
})
