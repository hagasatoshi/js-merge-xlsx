'use strict';

var express = require('express');
var path = require('path');
var bodyParser = require('body-parser');

var config = {
    jadeDir: './src/views',
    staticDir: 'public',
    indexFile: 'index'
};

var app = express();
app.set('views', path.join(__dirname, config.jadeDir));
app.set('view engine', 'jade');
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express['static'](config.staticDir));

app.get('/', function (req, res) {
    return res.render(config.indexFile);
});

var port = process.env.PORT || 3000;
app.set('port', port);
app.listen(app.get('port'), function () {
    return console.log('Listening on ' + port);
});

module.exports = app;