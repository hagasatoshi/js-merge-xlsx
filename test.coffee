JSZip = require 'JSZip'
Promise = require 'bluebird'
fs = Promise.promisifyAll(require('fs'))

SpreadSheet = require './build/spreadsheet'
spread = new SpreadSheet

fs.readFileAsync "./example/template/Blank.xlsx"
.then (data)->
  spread.initialize(data)
.then ()->
  spread.set_row 'Sheet1',8,['sample','user','pass','word']
  
  fs.writeFileAsync 'writetest.xlsx',spread.generate('nodebuffer')
.then ()->
  console.log 'Success'
.catch (err)->
  console.log err
