(function() {
  var SpreadSheet, _, angular;

  angular = require('angular');

  require('angular-bootstrap');

  require('angular-route');

  _ = require('underscore');

  SpreadSheet = require('./spreadsheet');

  angular.module('app', ['ui.bootstrap', 'ngRoute']).controller('extract_controller', function($scope, $http) {
    return $scope.save_template = function() {
      return $http({
        url: '/template/CustomField.xlsx',
        responseType: 'arraybuffer'
      }).success(function(data, status, headers, config) {
        var spread;
        spread = new SpreadSheet;
        return spread.initialize(data).then(function() {
          var excel;
          spread.copy_sheet('base', 'copied');
          excel = spread.generate('blob');
          saveAs(excel);
          return console.log("Success");
        });
      }).error(function(data, status, headers, config) {
        return console.log("Error");
      });
    };
  });

}).call(this);
