JSZip = require 'JSZip'
Promise = require 'bluebird'
fs = Promise.promisifyAll(require('fs'))

SpreadSheet = require './build/spreadsheet'
spread = new SpreadSheet

fs.readFileAsync "./example/template/CustomField.xlsx"
.then (data)->
  spread.initialize(data)
.then ()->
  spread.copy_sheet 'base','copied'
  spread.delete_sheet 'base'
  
  fs.writeFileAsync 'writetest.xlsx',spread.generate('nodebuffer')
.then ()->
  console.log 'Success'
.catch (err)->
  console.log err
