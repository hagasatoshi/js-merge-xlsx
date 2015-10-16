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
    data: readYamlAsync('./data/data1.yml'),                        //Load single data
        bulkData: readYamlAsync('./data/data2.yml'),                //Load array data
        excelMerge: new ExcelMerge().load(new JSZip(excelTemplate)) //Initialize ExcelMerge object
    });
}).then(({data, bulkData, excelMerge})=>{

    //add name property for ExcelMerge#bulkMergeMultiFile()
    let bulkData1 = _.map(bulkData, (e,index)=> ({name:`file${index+1}.xlsx`, data:e}));

    //add name property for ExcelMerge#bulkMergeMultiSheet()
    let bulkData2 = _.map(bulkData, (e,index)=> ({name:`example${index+1}`, data:e}));

    //Execute merge
    return Promise.props({
        excel1: excelMerge.merge(data),
        excel2: excelMerge.bulkMergeMultiFile(bulkData1),
        excel3: excelMerge.bulkMergeMultiSheet(bulkData2)
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