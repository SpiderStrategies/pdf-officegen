const _ = require('lodash')

const fs = require('fs')
const path = require('path')
const os = require('os')

const {info} = require('./logger')

module.exports = {

  getStagingDirectory (stagingDir) {
    return new Promise((resolve, reject) => {
      if (stagingDir) {
        fs.stat(stagingDir, (err, s) => {
          if (err || !s.isDirectory()) {
            info(this.options.jobId, 'staging directory:', stagingDir, 'does not exist, creating a new one')
            return this._createTempStagingDirectory()
          } else {
            this._createImageDirectory(stagingDir, reject, resolve)
          }
        })
      } else {
        resolve(this._createTempStagingDirectory())
      }
    })
  },

  _createTempStagingDirectory () {
    return new Promise((resolve, reject) => {
      fs.mkdtemp(path.join(os.tmpdir(), 'pdf_ppt_'), (err, folder) => {
        if (err) reject(err)
        this._createImageDirectory(folder, reject, resolve)
      })
    })
  },

  _createImageDirectory (folder, reject, resolve) {
    fs.mkdir(path.resolve(folder, 'img'), (err) => {
      if (err) reject(err)
      fs.mkdir(path.resolve(folder, 'pdf'), (err) => {
        if (err) reject(err)
        resolve(folder)
      })
    })
  },

  sortPages (imageFiles) {
    // Example: /var/folders/dr/f1q4znd96xv8wp82y4cfgg700000gn/T/pdf_ppt_5tz0dw/img/img-5-10.png
    // File = 5, Page = 10
    const rex = /.*(img|pdf)-(\d*)-(\d*).*/
    return imageFiles.sort((a, b) => {
      let aGrps = rex.exec(a)
      let bGrps = rex.exec(b)
      // PDF File Sequence + Page Sequence Comparison
      const fileGrp = 2
      const pageGrp = 3
      let fileComp = aGrps[fileGrp] - bGrps[fileGrp]
      if (fileComp === 0) {
        return aGrps[pageGrp] - bGrps[pageGrp]
      }
      return fileComp
    })
  },

  nowInMillis () {
    return Date.now() // process.hrtime()[1] / 1000000
  },

  elapsed (start) {
    return this.nowInMillis() - start
  }
}
