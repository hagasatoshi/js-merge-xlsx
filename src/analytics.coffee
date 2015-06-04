jsforce = require 'jsforce'
Promise = require 'bluebird'
_ = require 'underscore'

class Analytics
  initialize: (username, password, loginUrl)->
    loginUrl = if loginUrl then loginUrl else 'https://login.salesforce.com'
    @conn = new jsforce.Connection(loginUrl : loginUrl)
    @conn.login(username, password)
    
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
      return report_data

module.exports = Analytics
