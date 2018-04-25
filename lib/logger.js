const debug = require('debug')
const pdfLogger = debug('pdfppt:app')
const debugLogger = debug('pdfppt:debug')
pdfLogger.log = console.log.bind(console)
debugLogger.log = console.log.bind(console)

module.exports.info = pdfLogger
module.exports.debug = debugLogger
