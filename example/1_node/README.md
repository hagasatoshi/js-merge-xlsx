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
const readYamlAsync = Promise.promisify(require('read-yaml'));
const fs = Promise.promisifyAll(require('fs'));
const _ = require('underscore');
const {merge, bulkMergeToFiles, bulkMergeToSheets}
    = require('js-merge-xlsx');

Promise.props({
    templateObj: fs.readFileAsync('./template/Template.xlsx'),
    data: readYamlAsync('./data/data1.yml'),
    bulkData: readYamlAsync('./data/data2.yml')
}).then(({templateObj, data, bulkData}) => {

    let bulkData1 = _.map(bulkData, (e, index) =>{
        return {name: `file${index+1}.xlsx`, data: e};
    });

    let bulkData2 = _.map(bulkData, (e, index) => {
        return {name: `example${index+1}`, data: e};
    });

    return Promise.props({
        excel1: merge(templateObj, data),
        excel2: bulkMergeToFiles(templateObj, bulkData1),
        excel3: bulkMergeToSheets(templateObj, bulkData2)
    });
}).then(({excel1, excel2, excel3}) => {
    return Promise.all([
        fs.writeFileAsync('example1.xlsx', excel1),
        fs.writeFileAsync('example2.zip', excel2),
        fs.writeFileAsync('example3.xlsx', excel3)
    ]);
}).catch((err) => {
    console.error(err);
});
```
