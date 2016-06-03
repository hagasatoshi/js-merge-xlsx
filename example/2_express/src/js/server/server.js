const express = require('express');
const path = require('path');
const bodyParser = require('body-parser');

const config = {
    jadeDir: './src/views',
    staticDir: 'public',
    indexFile: 'index'
};

let app = express();
app.set('views', path.join(__dirname, config.jadeDir));
app.set('view engine', 'jade');
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({extended: true}));
app.use(express.static(config.staticDir));

app.get('/', (req, res) => res.render(config.indexFile));

var port = process.env.PORT || 3000;
app.set('port', port);
app.listen(app.get('port'), () => console.log(`Listening on ${port}`));

module.exports = app;
