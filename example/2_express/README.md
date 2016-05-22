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

- ExcelMerge#merge() : single merge.
- ExcelMerge#bulkMergeMultiFile() : bulk merge as 'multiple file'.
- ExcelMerge#bulkMergeMultiSheet() : bulk merge as 'multiple sheet'.
  
# Source  
Angular's controller(ES6 syntax)
```JavaScript
const Promise = require('bluebird');
const {merge, bulkMergeToFiles, bulkMergeToSheets} = require('js-merge-xlsx');
const JSZip = require('jszip');
const _ = require('underscore');

module.exports = ($scope, $http) => {

    $scope.merge = () => {
        Promise.props({
            template: $http.get('/template/Template.xlsx', {responseType: 'arraybuffer'}),
            data:     $http.get('/data/data1.json')
        }).then(({template, data}) => {

            //FileSaver#saveAs()
            saveAs(merge(template, data), 'example.xlsx');
        }).catch((err) => {
            console.log(err);
        });
    };

    $scope.bulkMergeToFiles = () => {
        Promise.props({
            template: $http.get('/template/Template.xlsx', {responseType: 'arraybuffer'}),
            data:     $http.get('/data/data2.json')
        }).then(({template, data}) => {

            data = _.map(data.data, (e,index) => {
                return {name: `file${(index+1)}.xlsx`, data: e};
            });
            //FileSaver#saveAs()
            saveAs(bulkMergeToFiles(template, data), 'example.zip');
        }).catch((err) => {
            console.log(err);
        });
    };

    $scope.bulkMergeToSheets = ()=>{
        Promise.props({
            template: $http.get('/template/Template.xlsx', {responseType: 'arraybuffer'}),
            data:     $http.get('/data/data2.json')
        }).then(({template, data}) => {

            data = _.map(data.data, (e,index) => {
                return {name: `sample${(index+1)}`, data: e};
            });

            //bulkMergeToSheets() is called asyc by returning Promise(bluebird) instance.
            return bulkMergeToSheets(template, data);
        }).then((excel) => {

            saveAs(excel,'example.xlsx');
        }).catch((err) => {
            console.log(err);
        });
    };
};
```
