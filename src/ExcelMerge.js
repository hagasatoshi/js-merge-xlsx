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
        return this.sheetHelper.simpleMerge(bindData);
    }

    bulkMergeMultiFile(bindDataArray){
        return this.sheetHelper.bulkMergeMultiFile(bindDataArray);
    }

    bulkMergeMultiSheet(bindDataArray){
        return this.sheetHelper.bulkMergeMultiSheet(bindDataArray);
    }

    variables(){
        return this.sheetHelper.excel.variables();
    }
}

module.exports = ExcelMerge;
