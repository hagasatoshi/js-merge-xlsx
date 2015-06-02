angular = require 'angular'
require 'angular-bootstrap'
require 'angular-route'
_ = require 'underscore'
SpreadSheet = require './spreadsheet'

angular.module 'app',[
  'ui.bootstrap',
  'ngRoute'
]
.controller 'extract_controller', ($scope, $http)->
  $scope.save_template = ()->
    $http {url:'/template/CustomField.xlsx',responseType: 'arraybuffer'}
    .success (data, status, headers, config)->
      spread = new SpreadSheet
      spread.initialize(data)
      .then ()->
        spread.copy_sheet 'base','copied'
        excel = spread.generate('blob')
        saveAs(excel)
        console.log "Success"
    .error (data, status, headers, config)->
      console.log "Error"