/**
 * ExcelMerge
 * top level api class for js-merge-xlsx
 * @author Satoshi Haga
 * @date 2015/09/30
 */

const Promise = require('bluebird');
const _ = require('underscore');
const JSZip = require('jszip');
const SheetHelper = require('./lib/sheetHelper');
const isNode = require('detect-node');

const output_buffer = {type: (isNode?'nodebuffer':'blob'), compression:"DEFLATE"};
const SINGLE_DATA = 'SINGLE_DATA';
const MULTI_FILE = 'MULTI_FILE';
const MULTI_SHEET = 'MULTI_SHEET';

class ExcelMerge{

    /**
     * constructor
     */
    constructor(){
        this.sheetHelper = new SheetHelper();
    }

    /**
     * load
     * @param {Object} excel JsZip object including MS-Excel file
     * @return {Promise} Promise instance including this
     */
    load(excel){

        //validation
        if(!(excel instanceof JSZip)){
            return Promise.reject('First parameter must be JSZip instance including MS-Excel data');
        }

        return this.sheetHelper.load(excel).then(()=>this);
    }

    /**
     * mergeByType
     * @param {String} mergeType
     * @param {Object} bindData binding data
     * @return {Promise} Promise instance including MS-Excel data. data-format is determined by jszip_option
     */
    mergeByType(mergeType, bindData){
        switch (mergeType){
            case SINGLE_DATA :
                return this.merge(bindData);
            case MULTI_FILE :
                return this.bulkMergeMultiFile(bindData);
            case MULTI_SHEET :
                return this.bulkMergeMultiSheet(bindData);
            default :
                return Promise.reject('Invalid parameter : mergeType');
        }

    }

    /**
     * merge
     * @param {Object} bindData binding data
     * @return {Promise} Promise instance including MS-Excel data. data-format is determined by jszip_option
     */
    merge(bindData){

        //validation
        if(!bindData){
            return Promise.reject('merge() must has parameter');
        }

        return this.sheetHelper.simpleMerge(bindData);
    }

    /**
     * bulkMergeMultiFile
     * @param {Array} bindDataArray including data{name: file's name, data: binding-data}
     * @return {Promise} Promise instance including MS-Excel data.
     */
    bulkMergeMultiFile(bindDataArray){

        //validation
        if(!bindDataArray){
            return Promise.reject('bulkMergeMultiFile() must has parameter');
        }
        return this.sheetHelper.bulkMergeMultiFile(bindDataArray);
    }

    /**
     * bulkMergeMultiSheet
     * @param {Array} bindDataArray including data{name: file's name, data: binding-data}
     * @return {Promise} Promise instance including MS-Excel data.
     */
    bulkMergeMultiSheet(bindDataArray){

        //validation
        if(!bindDataArray || !_.isArray(bindDataArray)) {
            return Promise.reject('bulkMergeMultiSheet() must has array as parameter');
        }

        _.each(bindDataArray, ({name,data})=>this.sheetHelper.addSheetBindingData(name,data));
        return this.sheetHelper.deleteTemplateSheet()
            .focusOnFirstSheet()
            .generate(output_buffer);
    }

    /**
     * variables
     * @return {Array}
     */
    variables(){
        return this.sheetHelper.templateVariables();
    }
}

//Exports
module.exports = ExcelMerge;
