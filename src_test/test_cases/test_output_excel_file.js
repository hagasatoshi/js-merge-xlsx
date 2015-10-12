/**
 * * test_output_excel_file.js
 * * Test code for spreadsheet
 * * @author Satoshi Haga
 * * @date 2015/10/11
 **/
var path = require('path');
var cwd = path.resolve('');
var assert = require('assert');
var JSZip = require('jszip');
var ExcelMerge = require(`${cwd}/excelmerge`);
var SpreadSheet = require(`${cwd}/lib/spreadsheet`);
require(cwd+'/lib/underscore_mixin');
var Promise = require('bluebird');
var readYamlAsync = Promise.promisify(require('read-yaml'));
var fs = Promise.promisifyAll(require('fs'));
var _ = require('underscore');

var EXCEL_OUTPUT_TYPE = {
    SINGLE : 0,
    BULK_MULTIPLE_FILE : 1,
    BULK_MULTIPLE_SHEET : 2
};

class Utility{

    output(templateName, inputFileName, outputType, outputFileName){
        return fs.readFileAsync(`${__dirname}/../templates/${templateName}`)
        .then((excelTemplate)=>{
            return Promise.props({
                renderingData: readYamlAsync(`${__dirname}/../input/${inputFileName}`),     //Load single data
                merge: new ExcelMerge().load(new JSZip(excelTemplate)) //Initialize ExcelMerge object
            });
        }).then(({renderingData,merge})=>{
            let dataArray = [];
            switch(outputType){
                case EXCEL_OUTPUT_TYPE.SINGLE:
                    return merge.render(renderingData);
                    break;

                case EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_FILE:
                    _.each(renderingData, (data,index)=> dataArray.push({name:`file${index+1}.xlsx`, data:data}));
                    return merge.bulkRenderMultiFile(dataArray);
                    break;

                case EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_SHEET:
                    _.each(renderingData, (data,index)=> dataArray.push({name:`example${index+1}`, data:data}));
                    return merge.bulkRenderMultiSheet(dataArray);
                    break;
            }
        }).then(
            outputData=>fs.writeFileAsync(`${__dirname}/../output/${outputFileName}`, outputData)
        ).then(
            ()=>assert(true)
        ).catch((err)=>{
            console.error(new Error(err).stack);
            assert(false);
        });
    }
}

module.exports = Utility;
