const path = require('path');
const bodyParser = require('body-parser');
const app = require('express')();

const config = {
    jadeFiles: './src/views',
    publicDir: './public',
    indexFile: 'index'
};

app.set('views', config.jadeFiles);
app.set('view engine', 'jade');
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({extended: true}));
app.use(express.static(config.publicDir));

app.get('/', (req, res)=> res.render(config.indexFile));

let port = process.env.PORT || 3000;
app.set('port', port);
app.listen(app.get('port'), ()=>console.log('Listening on ' + port));

module.exports = app;
