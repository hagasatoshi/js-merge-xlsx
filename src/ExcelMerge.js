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
     * * @return {Promise} Promise instance including this
     **/
    load(excel){

        //validation
        if(!(excel instanceof JSZip)){
            return Promise.reject('First parameter must be JSZip instance including MS-Excel data');
        }

        return this.spreadsheet.load(excel).then(()=>this);
    }

    /**
     * * merge
     * * @param {Object} bindData binding data
     * * @return {Promise} Promise instance including MS-Excel data. data-format is determined by jszip_option
     **/
    merge(bindData){

        //validation
        if(!bindData){
            return Promise.reject('merge() must has parameter');
        }

        return this.spreadsheet.simpleMerge(bindData);
    }

    /**
     * * bulkMergeMultiFile
     * * @param {Array} bindDataArray including data{name: file's name, data: binding-data}
     * * @return {Promise} Promise instance including MS-Excel data.
     **/
    bulkMergeMultiFile(bindDataArray){

        //validation
        if(!bindDataArray){
            return Promise.reject('bulkMergeMultiFile() must has parameter');
        }
        return this.spreadsheet.bulkMergeMultiFile(bindDataArray);
    }

    /**
     * * bulkMergeMultiSheet
     * * @param {Array} bindDataArray including data{name: file's name, data: binding-data}
     * * @return {Promise} Promise instance including MS-Excel data.
     **/
    bulkMergeMultiSheet(bindDataArray){

        //validation
        if(!bindDataArray || !_.isArray(bindDataArray)) {
            return Promise.reject('bulkMergeMultiSheet() must has array as parameter');
        }

        _.each(bindDataArray, ({name,data})=>this.spreadsheet.addSheetBindingData(name,data));
        return this.spreadsheet.deleteTemplateSheet()
            .focusOnFirstSheet()
            .generate(output_buffer);
    }
}

//Exports
module.exports = ExcelMerge;
