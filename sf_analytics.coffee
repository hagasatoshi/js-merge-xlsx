jsforce = require 'jsforce'
Promise = require 'bluebird'
fs = Promise.promisifyAll(require('fs'))
_ = require 'underscore'
require('dotenv').load()

Analytics = require './build/analytics'
SpreadSheet = require './build/spreadsheet'

analytics = new Analytics()
spread = new SpreadSheet()

fs.readFileAsync "./example/template/Blank.xlsx"
.then (data)->
  Promise.all [
    analytics.initialize(process.env.SALESFORCE_USERNAME, process.env.SALESFORCE_PASSWORD),
    spread.initialize data
  ]
.then ()->
  #analytics.report_data '00O10000004NerC'
  analytics.report_data '00O10000003uj5S'
.then (report_data)->
  _.each report_data, (row,index)->
    spread.set_row 'Sheet1',index,row
  fs.writeFileAsync 'analytics.xlsx',spread.generate('nodebuffer')
.then ()->
  console.log 'Success'
.catch (err)->
  console.log err
