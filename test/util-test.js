import {test} from 'ava'
import util from '../lib/util'

test('sort pages', t => {
  // 1-100 to ensure both the doc and page are sorted independently
  const images = ['img/img-21-9.png', 'img/img-1-1.png', 'img/img-1-100.png']
  const sorted = util.sortPages(images)
  t.deepEqual(sorted, ['img/img-1-1.png', 'img/img-1-100.png', 'img/img-21-9.png'])
})
