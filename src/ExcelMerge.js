/**
 * * ExcelMerge
 * * top level api class for js-merge-xlsx
 * * @author Satoshi Haga
 * * @date 2015/09/30
 **/

var Mustache = require('mustache');
var Promise = require('bluebird');
var _ = require('underscore');
var JSZip = require('jszip');
var SpreadSheet = require('./lib/spreadsheet');
var isNode = require('detect-node');
const output_buffer = {type: (isNode?'nodebuffer':'blob'), compression:"DEFLATE"};

class ExcelMerge{

    /**
     * * constructor
     * *
     **/
    constructor(){
        this.spreadsheet = new SpreadSheet();
    }

    /**
     * * load
     * * @param {Object} excel JsZip object including MS-Excel file
     * * @param {Object} option option parameter
     * * @return {Promise} Promise instance including this
     **/
    load(excel, option){
        return this.spreadsheet.load(excel, option).then(()=>this);
    }

    /**
     * * render
     * * @param {Object} bind_data binding data
     * * @returns {Object} rendered MS-Excel data. data-format is determined by jszip_option
     **/
    render(bind_data){
        return this.spreadsheet.simpleRender(bind_data);
    }

    /**
     * * bulkRenderMultiFile
     * * @param {Array} bind_data_array including data{name: file's name, data: binding-data}
     * * @returns {Object} rendered MS-Excel data.
     **/
    bulkRenderMultiFile(bind_data_array){
        return this.spreadsheet.bulkRenderMultiFile(bind_data_array);
    }

    /**
     * * bulkRenderMultiSheet
     * * @param {Array} bind_data_array including data{name: file's name, data: binding-data}
     * * @returns {Object} rendered MS-Excel data. data-format is determined by jszip_option
     **/
    bulkRenderMultiSheet(bind_data_array){
        return bind_data_array.reduce(
            (promise, {name, data})=>
                promise.then((prior)=>{
                    return this.spreadsheet.addSheetBindingData(name,data);
                })
            , Promise.resolve()
        ).then(()=>{
            return this.spreadsheet.deleteTemplateSheet()
                .forcusOnFirstSheet()
                .generate(output_buffer);
        }).catch((err)=>{
            console.error(new Error(err).stack);
            Promise.reject();
        });
    }

}

//Exports
module.exports = ExcelMerge;
