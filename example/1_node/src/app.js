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