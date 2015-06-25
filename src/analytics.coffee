jsforce = require 'jsforce'
Promise = require 'bluebird'
_ = require 'underscore'
require './mixin'

class Analytics
  initialize: (conn)->
    @conn = conn

  report_data: (report_id)->
    @conn.analytics.report report_id
    .execute { details: true }
    .then (result)->
      report_data = []
      _.each result.factMap["T!T"].rows, (row)->
        cells = []
        _.each row.dataCells, (cell)->
          cells.push cell.label
        report_data.push cells
      console.log report_data
      return report_data
    .fail (err)->
      console.log err

module.exports = Analytics
