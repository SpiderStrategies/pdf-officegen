import {test} from 'ava'
import Engine from '../lib/engine'

let engine

test.beforeEach(t => {
  engine = new Engine({engine: ''})
})

test('Convert options do not allow density > 300', t => {
  engine = new Engine({ convertOptions: { 'density': 600 } })
  const o = engine._getConvertOptions({ convertOptions: { 'density': 600 } })
  t.is(o['density'], 300)
})

test('Convert options use density=72 if none is set', t => {
  const o = engine._getConvertOptions()
  t.is(o['density'], 72)
})

test('_getLastImageFiles with different files', t => {
  // Only precondition should be that the images are ordered for each file
  const images = ['img/img-1-1.png', 'img/img-1-100.png', 'img/img-21-9.png', 'img/img-2-1.png']
  const lastImages = engine._getLastImageFiles(images)
  console.log(lastImages)
  t.deepEqual(lastImages, ['img/img-1-100.png', 'img/img-21-9.png', 'img/img-2-1.png'])
})

test('_getLastImageFiles with same file', t => {
  // Only precondition should be that the images are ordered for each file
  const images = ['img/img-1-0.png', 'img/img-1-1.png', 'img/img-1-2.png']
  const lastImages = engine._getLastImageFiles(images)
  console.log(lastImages)
  t.deepEqual(lastImages, ['img/img-1-2.png'])
})
