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
js-merge-xlsx supports Promises/A+([bluebird](https://github.com/petkaantonov/bluebird)). So, it is called basically in Promise-chain.  
Example(ES6 syntax)  
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
        data: readYamlAsync('./data/data1.yml'),
        bulkData: readYamlAsync('./data/data2.yml'),
        excelMerge1: new ExcelMerge().load(new JSZip(excelTemplate)),
        excelMerge2: new ExcelMerge().load(new JSZip(excelTemplate)),
        excelMerge3: new ExcelMerge().load(new JSZip(excelTemplate))
    });
}).then(({data, bulkData, excelMerge1, excelMerge2, excelMerge3})=>{

    //add name property for ExcelMerge#bulkMergeMultiFile()
    let bulkData1 = _.map(bulkData, (e,index)=> ({name:`file${index+1}.xlsx`, data:e}));

    //add name property for ExcelMerge#bulkMergeMultiSheet()
    let bulkData2 = _.map(bulkData, (e,index)=> ({name:`example${index+1}`, data:e}));

    //Execute merge
    return Promise.props({
        excel1: excelMerge1.merge(data),
        excel2: excelMerge2.bulkMergeMultiFile(bulkData1),
        excel3: excelMerge3.bulkMergeMultiSheet(bulkData2)
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
