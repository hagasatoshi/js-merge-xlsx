# Example on Node.js  
Here is example on Node.js.
  
# git clone
```bash
git clone git@github.com:hagasatoshi/js-merge-xlsx.git
cd js-merge-xlsx/example/1_node
```
# Build
```bash
npm install
gulp
```
# Execute
```bash
node app.js
```
The following 3 files are created.
- example1.xlsx : single printing.
- example2.zip : bulk printing. This zip file contains 3 files.
- example3.xlsx : bulk printing. This file has 3 Excel-sheets.
  
# Source  
app.js
```JavaScript
const Promise = require('bluebird');
const readYamlSync = require('read-data').yaml.sync;
const fs = Promise.promisifyAll(require('fs'));
const _ = require('underscore');
const {merge, bulkMergeToFiles, bulkMergeToSheets}
    = require('js-merge-xlsx');

const config = {
    template:   './template/Template.xlsx',
    singleData: './data/data1.yml',
    arrayData:  './data/data2.yml'
};

const readData = () => {
    let templateObj = fs.readFileSync(config.template);
    let data  = readYamlSync(config.singleData);
    let bulkData = readYamlSync(config.arrayData);

    return {
        templateObj: templateObj,
        data: data,
        bulkData1: _.map(bulkData, (e, index) => {
            return {name: `file${index + 1}.xlsx`, data: e};
        }),
        bulkData2: _.map(bulkData, (e, index) => {
            return {name: `example${index + 1}`, data: e};
        })
    };
};

//Start
let {templateObj, data, bulkData1, bulkData2} = readData();

//example of merge()
fs.writeFileSync('example1.xlsx',  merge(templateObj, data));

//example of bulkMergeToFiles()
fs.writeFileSync(
    'example2.zip',
    bulkMergeToFiles(templateObj, bulkData1)
);

//example of bulkMergeToSheets()
//this method is called async by returning Promise(bluebird) instance.
bulkMergeToSheets(templateObj, bulkData2)
.then((excel) => {
    fs.writeFileSync('example3.xlsx', excel);
});
```
