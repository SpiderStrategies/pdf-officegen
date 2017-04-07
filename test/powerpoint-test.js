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

test('Convert options do not allow density > 300', t => {
  const o = p._getConvertOptions({ convertOptions: { 'density': 600 } })
  t.is(o['density'], 300)
})

test('Convert options use density=72 if none is set', t => {
  const o = p._getConvertOptions({})
  t.is(o['density'], 72)
})

test('sort images', t => {
  // 1-100 to ensure both the doc and page are sorted independently
  const images = ['img/img-21-9.png', 'img/img-1-1.png', 'img/img-1-100.png']
  const sorted = p._sortImages(images)
  t.deepEqual(sorted, ['img/img-1-1.png', 'img/img-1-100.png', 'img/img-21-9.png'])
})

test('_getLastImageFiles with different files', t => {
  // Only precondition should be that the images are ordered for each file
  const images = ['img/img-1-1.png',
                  'img/img-1-100.png',
                  'img/img-21-9.png',
                  'img/img-2-1.png']
  const lastImages = p._getLastImageFiles(images)
  console.log(lastImages)
  t.deepEqual(lastImages, ['img/img-1-100.png', 'img/img-21-9.png', 'img/img-2-1.png'])
})

test('_getLastImageFiles with same file', t => {
  // Only precondition should be that the images are ordered for each file
  const images = ['img/img-1-0.png',
                  'img/img-1-1.png',
                  'img/img-1-2.png']
  const lastImages = p._getLastImageFiles(images)
  console.log(lastImages)
  t.deepEqual(lastImages, ['img/img-1-2.png'])
})

