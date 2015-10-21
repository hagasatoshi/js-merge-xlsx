/**
 * * server.js
 * * Express server-side process
 * * @author Satoshi Haga
 * * @date 2015/10/05
 **/
'use strict';

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { 'default': obj }; }

var _express = require('express');

var _express2 = _interopRequireDefault(_express);

var _path = require('path');

var _path2 = _interopRequireDefault(_path);

var _bodyParser = require('body-parser');

var _bodyParser2 = _interopRequireDefault(_bodyParser);

var app = (0, _express2['default'])();
app.set('views', _path2['default'].join(__dirname, './src/views'));
app.set('view engine', 'jade');
app.use(_bodyParser2['default'].json());
app.use(_bodyParser2['default'].urlencoded({ extended: true }));
app.use(_express2['default']['static']('public'));

app.get('/', function (req, res) {
  return res.render('index');
});

var port = process.env.PORT || 3000;
app.set('port', port);
app.listen(app.get('port'), function () {
  return console.log('Listening on ' + port);
});

module.exports = app;