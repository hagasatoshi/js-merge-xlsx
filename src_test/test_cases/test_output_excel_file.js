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

    output(template_name, input_file_name, output_type, output_file_name){
        return fs.readFileAsync(`${__dirname}/../templates/${template_name}`)
        .then((excel_template)=>{
            return Promise.props({
                rendering_data: readYamlAsync(`${__dirname}/../input/${input_file_name}`),     //Load single data
                merge: new ExcelMerge().load(new JSZip(excel_template)) //Initialize ExcelMerge object
            });
        }).then((result)=>{
            //ExcelMerge object
            let merge =  result.merge;

            //rendering data
            let rendering_data;
            if(output_type === EXCEL_OUTPUT_TYPE.SINGLE){
                rendering_data = result.rendering_data;
                return merge.render(rendering_data);
            }else if(output_type === EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_FILE){
                rendering_data = [];
                _.each(result.rendering_data, (data,index)=>{
                    rendering_data.push({name:`file${index+1}.xlsx`, data:data});
                });
                return merge.bulkRenderMultiFile(rendering_data);
            }else if(output_type === EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_SHEET){
                rendering_data = [];
                _.each(result.rendering_data, (data,index)=>{
                    rendering_data.push({name:`example${index+1}`, data:data});
                });
                return merge.bulkRenderMultiSheet(rendering_data);
            }
        }).then((output_data)=>{
            return fs.writeFileAsync(`${__dirname}/../output/${output_file_name}`, output_data);
        }).then(()=>{
            assert(true);
        }).catch((err)=>{
            console.error(new Error(err).stack);
            assert(false);
        });
    }
}

module.exports = Utility;
