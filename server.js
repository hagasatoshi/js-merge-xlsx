var express = require('express');
var path = require('path');
var session = require('express-session');
var bodyParser = require('body-parser');
var app = express();
app.set('views', path.join(__dirname, './src/views'));
app.set('view engine', 'jade');
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({extended: true}));
app.use(express.static('public'));

var ERROR_MESSAGE = '予期せぬエラーが発生しました';

app.get('/index', function(req, res) {
    res.render('index');
});

var port = process.env.PORT || 3000;
app.set('port', port);
var server = app.listen(app.get('port'), function(){
    console.log("Listening on " + port);
});

module.exports = app;
