jsforce = require 'jsforce'
Promise = require 'bluebird'
_ = require 'underscore'
require('dotenv').load()

Analytics = require './build/analytics'
SpreadSheet = require './build/spreadsheet'

analytics = new Analytics();

analytics.initialize process.env.SALESFORCE_USERNAME, process.env.SALESFORCE_PASSWORD
.then ()->
  analytics.report_data '00O10000003uj5S'
.then (report_data)->
  console.log report_data
.fail (err)->
  console.log err
