/**
 * * app.js
 * * client-side process
 * * @author Satoshi Haga
 * * @date 2015/10/06
 **/

import angular from 'angular'
import 'angular-bootstrap'

angular.module('app',[
    'ui.bootstrap',
    require('./controllers').name
]).config(($locationProvider,$httpProvider)=>{
    if(!$httpProvider.defaults.headers.get)
        $httpProvider.defaults.headers.get = {};
    $httpProvider.defaults.headers.get['If-Modified-Since'] = '0';
    $locationProvider.html5Mode({enabled: true,requireBase: false});
});