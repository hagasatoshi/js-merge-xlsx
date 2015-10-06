# example on Node.js  
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

import ExcelMerge from 'js-merge-xlsx'
import Promise from 'bluebird'
var readYamlAsync = Promise.promisify(require('read-yaml'));
var fs = Promise.promisifyAll(require('fs'));
import JSZip from 'jszip'
import _ from 'underscore'

//Load Template
fs.readFileAsync('./template/Template.xlsx')
.then((excel_template)=>{
    return Promise.props({
        rendering_data1: readYamlAsync('./data/data1.yml'),     //Load single data
        rendering_data2: readYamlAsync('./data/data2.yml'),     //Load array data
        merge: new ExcelMerge().load(new JSZip(excel_template)) //Initialize ExcelMerge object
    });
}).then((result)=>{
    //Single-printing
    let rendering_data1 = result.rendering_data1;

    //Bulk-printing as 'multiple files'
    let rendering_data2 = [];
    _.each(result.rendering_data2, (data,index)=>{
        rendering_data2.push({name:'file'+(index+1)+'.xlsx', data:data});
    });

    //Bulk-printing as 'multiple sheets'
    let rendering_data3 = [];
    _.each(result.rendering_data2, (data,index)=>{
        rendering_data3.push({name:'example'+(index+1), data:data});
    });

    //ExcelMerge object
    let merge =  result.merge;

    //Execute rendering
    return Promise.props({
        excel_data1: merge.render(rendering_data1),
        excel_data2: merge.bulk_render_multi_file(rendering_data2),
        excel_data3: merge.bulk_render_multi_sheet(rendering_data3)
    });
}).then((result)=>{
    return Promise.all([
        fs.writeFileAsync('Example1.xlsx',result.excel_data1),
        fs.writeFileAsync('Example2.zip',result.excel_data2),
        fs.writeFileAsync('Example3.xlsx',result.excel_data3)
    ]);
}).catch((err)=>{
    console.error(new Error(err).stack);
});
```
