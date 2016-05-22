const angular = require('angular');
require('angular-bootstrap');

angular.module('app', [
    'ui.bootstrap',
    require('./controllers').name
]).config(($locationProvider, $httpProvider) => {

    let headers = $httpProvider.defaults.headers;
    headers.get = headers.get ? headers.get : {};

    headers.get['If-Modified-Since'] = '0';
    $locationProvider.html5Mode({
        enabled: true,requireBase: false
    });
});