import {test} from 'ava'
import Powerpoint from '../lib/powerpoint'

const p = new Powerpoint()

// Need to set GS Path for this to work
test.skip.cb('image is resized to fit on slide', t => {
  const options = {}
  p.convertPDFToPowerpoint('google-l.pdf', options, (err, result) => {
    if (err) {
      t.fail('Did not complete successfully')
      console.log(err)
    }
    t.end()
  })
})

test('Uses default slide dimensions if they are passed in undefined', t => {
  var dimensions = {
    height: undefined,
    width: undefined,
    type: undefined
  }
  let p = new Powerpoint({dimensions})
  t.is(p.options.dimensions.width, 800)
  t.is(p.options.dimensions.height, 600)
  t.is(p.options.dimensions.type, 'screen4x3')
})

test('Uses slide dimensions if they are passed as options', t => {
  var dimensions = {
    height: 1000
  }
  let p = new Powerpoint({dimensions})
  t.is(p.options.dimensions.width, 800)
  t.is(p.options.dimensions.height, 1000)
  t.is(p.options.dimensions.type, 'screen4x3')
})

test('sets cropLastImage to false for backwards compatibility', t => {
  let p = new Powerpoint()
  t.is(p.options.cropLastImage, false)
})

test('allows cropLastImage to be set in config', t => {
  let p = new Powerpoint({cropLastImage: true})
  t.is(p.options.cropLastImage, true)
})
