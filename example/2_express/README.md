# Example on web browser  
Here is example on web browser. Webpack(or browserify) empower you to use node-modules on web-browser as well. Also bluebird automatically casts thenable object, such as object returned by '$http.get()' and '$.get()', to trusted Promise. So, you can code in the same way as Node.js.  
  
# git clone
```bash
git clone git@github.com:hagasatoshi/js-merge-xlsx.git
cd js-merge-xlsx/example/2_express
```
# Build
```bash
npm install
bower install
gulp
```
# Start server
```bash
node server.js
```

You can access 'http://localhost:3000/' as follows.  
![Experss](https://raw.githubusercontent.com/hagasatoshi/js-merge-xlsx/master/image/express.png)  
  
Each button prints using js-merge-xlsx.

- ExcelMerge#render() : single printing.
- ExcelMerge#bulk_render_multi_file() : bulk printing as 'multiple file'.
- ExcelMerge#bulk_render_multi_sheet() : bulk printing as 'multiple sheet'.
  
# Source  
Angular's controller(ES6 syntax)
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
