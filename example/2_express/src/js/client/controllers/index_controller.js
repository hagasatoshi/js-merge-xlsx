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
     * * exampleRender
     * * example of ExcelMerge#render()
     */
    $scope.exampleRender = ()=>{
        Promise.resolve($http.get('/template/Template.xlsx', {responseType: "arraybuffer"}))
        .then((excelTemplate)=>{
            return Promise.props({
                data: $http.get('/data/data1.json'),
                excelMerge: new ExcelMerge().load(new JSZip(excelTemplate.data))
            });
        }).then(({data, excelMerge})=>{
            return excelMerge.render(data.data);
        }).then((excelData)=>{
            saveAs(excelData,'example.xlsx');   //FileSaver#saveAs()
        }).catch((err)=>{
            console.error(err);
        });
    };

    /**
     * * exampleBulkRenderMultiFile
     * * example of ExcelMerge#exampleBulkRenderMultiFile()
     */
    $scope.exampleBulkRenderMultiFile = ()=>{
        Promise.resolve($http.get('/template/Template.xlsx', {responseType: "arraybuffer"}))
        .then((excelTemplate)=>{
            return Promise.props({
                data: $http.get('/data/data2.json'),
                excelMerge: new ExcelMerge().load(new JSZip(excelTemplate.data))
            });
        }).then(({data, excelMerge})=>{
            data = _.map(data.data, (e,index)=>({name:`file${(index+1)}.xlsx`, data:e}));
            return excelMerge.bulkRenderMultiFile(data); //FileSaver#saveAs()
        }).then((zipData)=>{
            saveAs(zipData,'example.zip');
        }).catch((err)=>{
            console.error(err);
        });
    };

    /**
     * * exampleBulkRenderMultiSheet
     * * example of ExcelMerge#exampleBulkRenderMultiSheet()
     */
    $scope.exampleBulkRenderMultiSheet = ()=>{
        Promise.resolve($http.get('/template/Template.xlsx', {responseType: "arraybuffer"}))
        .then((excelTemplate)=>{
            return Promise.props({
                data: $http.get('/data/data2.json'),
                excelMerge: new ExcelMerge().load(new JSZip(excelTemplate.data))
            });
        }).then(({data, excelMerge})=>{
            data = _.map(data.data, (e,index)=>({name:`sample${(index+1)}`, data:e}));
            return excelMerge.bulkRenderMultiSheet(data);
        }).then((excelData)=>{
            saveAs(excelData,'example.xlsx');
        }).catch((err)=>{
            console.error(err);
        });
    };
};

module.exports = indexController;