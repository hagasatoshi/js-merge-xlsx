/**
 * * server.js
 * * Express server-side process
 * * @author Satoshi Haga
 * * @date 2015/10/05
 **/
import express from 'express'
import path from 'path'
import bodyParser from 'body-parser'
var app = express();
app.set('views', path.join(__dirname, './src/views'));
app.set('view engine', 'jade');
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({extended: true}));
app.use(express.static('public'));

app.get('/', (req, res)=> res.render('index'));

var port = process.env.PORT || 3000;
app.set('port', port);
app.listen(app.get('port'), ()=>console.log('Listening on ' + port));

module.exports = app;
