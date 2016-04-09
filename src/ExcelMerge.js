/**
 * ExcelMerge
 * top level api class for js-merge-xlsx
 * @author Satoshi Haga
 * @date 2015/09/30
 */

const Promise = require('bluebird');
const _ = require('underscore');
const Excel = require('./lib/Excel');
const SheetHelper = require('./lib/sheetHelper');
const isNode = require('detect-node');
const output_buffer = {type: (isNode?'nodebuffer':'blob'), compression:"DEFLATE"};
const SINGLE_DATA = 'SINGLE_DATA';
const MULTI_FILE = 'MULTI_FILE';
const MULTI_SHEET = 'MULTI_SHEET';

class ExcelMerge{

    constructor(){
        this.sheetHelper = new SheetHelper();
    }

    load(excel){
        if(!(excel instanceof Excel)){
            return Promise.reject('First parameter must be Excel instance including MS-Excel data');
        }
        return this.sheetHelper.load(excel).then(()=>this);
    }

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

    merge(bindData){
        if(!bindData){
            return Promise.reject('merge() must has parameter');
        }
        return this.sheetHelper.simpleMerge(bindData);
    }

    bulkMergeMultiFile(bindDataArray){
        if(!bindDataArray){
            return Promise.reject('bulkMergeMultiFile() must has parameter');
        }
        return this.sheetHelper.bulkMergeMultiFile(bindDataArray);
    }

    bulkMergeMultiSheet(bindDataArray){
        if(!bindDataArray || !_.isArray(bindDataArray)) {
            return Promise.reject('bulkMergeMultiSheet() must has array as parameter');
        }

        _.each(bindDataArray, ({name,data})=>this.sheetHelper.addSheetBindingData(name,data));
        return this.sheetHelper.deleteTemplateSheet()
            .generate(output_buffer);
    }

    variables(){
        return this.sheetHelper.excel.variables();
    }
}

module.exports = ExcelMerge;
