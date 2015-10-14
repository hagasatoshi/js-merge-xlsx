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
- Example1.xlsx : single printing.
- Example2.zip : bulk printing. This zip file contains 3 files.
- Example3.xlsx : bulk printing. This file has 3 Excel-sheets.
  
# Source  
app.js(ES6 syntax)
```JavaScript
/**
 * * app.js
 * * Example on Node.js
 * * @author Satoshi Haga
 * * @date 2015/09/30
 **/

var ExcelMerge = require('js-merge-xlsx');
var Promise = require('bluebird');
var readYamlAsync = Promise.promisify(require('read-yaml'));
var fs = Promise.promisifyAll(require('fs'));
var JSZip = require('jszip');
var _ = require('underscore');

//Load Template
fs.readFileAsync('./template/Template.xlsx')
.then((excelTemplate)=>{
    return Promise.props({
        data: readYamlAsync('./data/data1.yml'),                //Load single data
        bulkData: readYamlAsync('./data/data2.yml'),            //Load array data
        merge: new ExcelMerge().load(new JSZip(excelTemplate))  //Initialize ExcelMerge object
    });
}).then(({data, bulkData, merge})=>{

    //add name property for ExcelMerge#bulkRenderMultiFile()
    let bulkData1 = _.map(bulkData, (e,index)=> ({name:`file${index+1}.xlsx`, data:e}));

    //add name property for ExcelMerge#bulkRenderMultiSheet()
    let bulkData2 = _.map(bulkData, (e,index)=> ({name:`example${index+1}`, data:e}));

    //Execute rendering
    return Promise.props({
        excel1: merge.render(data),
        excel2: merge.bulkRenderMultiFile(bulkData1),
        excel3: merge.bulkRenderMultiSheet(bulkData2)
    });
}).then(({excel1, excel2, excel3})=>{
    return Promise.all([
        fs.writeFileAsync('example1.xlsx',excel1),
        fs.writeFileAsync('example2.zip',excel2),
        fs.writeFileAsync('example3.xlsx',excel3)
    ]);
}).catch((err)=>{
    console.error(new Error(err).stack);
});
```
