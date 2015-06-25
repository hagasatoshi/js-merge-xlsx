jsforce = require 'jsforce'
Promise = require 'bluebird'
_ = require 'underscore'
Analytics = require './build/analytics'
SpreadSheet = require './build/spreadsheet'

class AnalyticsReport
  ###
    initialize
    @param {arraybuffer} excel file data(openXML format)
    @param {jsforce.Connection} connection object of jsforce
  ###
  initialize: (excel, conn)->
    @analytics = new SpreadSheet()
    @analytics.initialize excel
    @conn = new SpreadSheet()
    @conn = @conn.initialize conn

  ###
    set_report
    @param {arraybuffer} excel file data(openXML format)
  ###
  set_report: (excel)->
    @conn = new SpreadSheet()
    @conn = @conn.initialize conn

  build_report: ()->

    