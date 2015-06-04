(function() {
  var JSZip, Promise, SpreadSheet, fs, spread;

  JSZip = require('JSZip');

  Promise = require('bluebird');

  fs = Promise.promisifyAll(require('fs'));

  SpreadSheet = require('./build/spreadsheet');

  spread = new SpreadSheet;

  fs.readFileAsync("./example/template/Blank.xlsx").then(function(data) {
    return spread.initialize(data);
  }).then(function() {
    spread.set_row('Sheet1', 8, ['sample', 'user', 'pass', 'word']);
    return fs.writeFileAsync('writetest.xlsx', spread.generate('nodebuffer'));
  }).then(function() {
    return console.log('Success');
  })["catch"](function(err) {
    return console.log(err);
  });

}).call(this);
