# js-merge-xlsx  [![Build Status](https://travis-ci.org/hagasatoshi/js-merge-xlsx.svg?branch=master)](https://travis-ci.org/hagasatoshi/js-merge-xlsx)
Minimum JavaScript-based template engine for MS-Excel. js-merge-xlsx empowers you to print JavaScript objects.

- Available for both web browser and Node.js .
- Bulk printing. It is possible to print array as 'multiple files'. 
- Bulk printing. It is possible to print array as 'multiple sheets'. 

Template  
![Template](https://raw.githubusercontent.com/hagasatoshi/js-merge-xlsx/master/image/before2.png)  
After printing  
![Rendered](https://raw.githubusercontent.com/hagasatoshi/js-merge-xlsx/master/image/after.png)  

# Install
```bash
npm install js-merge-xlsx
```

# Prepare template  
Prepare the template with bind-variables as mustache format {{}}.
![Template](https://raw.githubusercontent.com/hagasatoshi/js-merge-xlsx/master/image/before2.png)  
**Note**: Only string cell is supported. Please make sure that the format of cells having variables is String.  
![Note](https://raw.githubusercontent.com/hagasatoshi/js-merge-xlsx/master/image/cell_format.png)

# Node.js  
example  
```JavaScript
const Promise = require('bluebird');
const readYamlSync = require('read-data').yaml.sync;
const fs = Promise.promisifyAll(require('fs'));
const _ = require('underscore');
const {merge, bulkMergeToFiles, bulkMergeToSheets} = require('js-merge-xlsx');

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

Please check [example codes](https://github.com/hagasatoshi/js-merge-xlsx/tree/master/example/1_node) and [API](https://github.com/hagasatoshi/js-merge-xlsx/blob/master/API.md) for detail.

# Browser  
You can also use it on web browser by using webpack(browserify). 
Bluebird automatically casts thenable object, such as object returned by "$http.get()" or "$.get()", to trusted Promise. https://github.com/petkaantonov/bluebird/blob/master/API.md#promiseresolvedynamic-value---promise  
So, you can code in the same way as Node.js.    
Example(ES6 syntax)  
```JavaScript
/**
 * * indexController.js
 * * angular controller definition
 * * @author Satoshi Haga
 * * @date 2015/10/06
 **/

var Promise = require('bluebird');
var ExcelMerge = require('js-merge-xlsx');
var JSZip = require('jszip');
var _ = require('underscore');

var indexController = ($scope, $http)=>{

    /**
     * * exampleMerge
     * * example of ExcelMerge#merge()
     */
    $scope.exampleMerge = ()=>{
        Promise.resolve($http.get('/template/Template.xlsx', {responseType: "arraybuffer"}))
        .then((excelTemplate)=>{
            return Promise.props({
                data: $http.get('/data/data1.json'),
                excelMerge: new ExcelMerge().load(new JSZip(excelTemplate.data))
            });
        }).then(({data, excelMerge})=>{
            return excelMerge.merge(data.data);
        }).then((excelData)=>{
            saveAs(excelData,'example.xlsx');   //FileSaver#saveAs()
        }).catch((err)=>{
            console.error(err);
        });
    };

    /**
     * * exampleBulkMergeMultiFile
     * * example of ExcelMerge#bulkMergeMultiFile()
     */
    $scope.exampleBulkMergeMultiFile = ()=>{
        Promise.resolve($http.get('/template/Template.xlsx', {responseType: "arraybuffer"}))
        .then((excelTemplate)=>{
            return Promise.props({
                data: $http.get('/data/data2.json'),
                excelMerge: new ExcelMerge().load(new JSZip(excelTemplate.data))
            });
        }).then(({data, excelMerge})=>{
            data = _.map(data.data, (e,index)=>({name:`file${(index+1)}.xlsx`, data:e}));
            return excelMerge.bulkMergeMultiFile(data); //FileSaver#saveAs()
        }).then((zipData)=>{
            saveAs(zipData,'example.zip');
        }).catch((err)=>{
            console.error(err);
        });
    };

    /**
     * * exampleBulkMergeMultiSheet
     * * example of ExcelMerge#bulkMergeMultiSheet()
     */
    $scope.exampleBulkMergeMultiSheet = ()=>{
        Promise.resolve($http.get('/template/Template.xlsx', {responseType: "arraybuffer"}))
        .then((excelTemplate)=>{
            return Promise.props({
                data: $http.get('/data/data2.json'),
                excelMerge: new ExcelMerge().load(new JSZip(excelTemplate.data))
            });
        }).then(({data, excelMerge})=>{
            data = _.map(data.data, (e,index)=>({name:`sample${(index+1)}`, data:e}));
            return excelMerge.bulkMergeMultiSheet(data);
        }).then((excelData)=>{
            saveAs(excelData,'example.xlsx');
        }).catch((err)=>{
            console.error(err);
        });
    };
};

module.exports = indexController;
```

Please check [example codes](https://github.com/hagasatoshi/js-merge-xlsx/tree/master/example/2_express) and [API](https://github.com/hagasatoshi/js-merge-xlsx/blob/master/API.md) for detail.
