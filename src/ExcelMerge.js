const Promise = require('bluebird');
const _ = require('underscore');
require('./lib/underscore_mixin');

const Excel = require('./lib/Excel');
const WorkBookXml = require('./lib/WorkBookXml');
const WorkBookRels = require('./lib/WorkBookRels');
const SheetXmls = require('./lib/SheetXmls');
const SharedStrings = require('./lib/SharedStrings');

const isNode = require('detect-node');
const config = {
    compression: "DEFLATE",
    buffer_type_output: (isNode ? 'nodebuffer' : 'blob'),
    buffer_type_jszip: (isNode ? 'nodebuffer' : 'arraybuffer')
};

class ExcelMerge{

    /**
     * load
     * @param {Excel} excel
     * @return {Promise} this
     */
    load(excel){
        this.excel = excel;
        return Promise.props({
            sharedstrings: excel.parseSharedStrings(),
            workbookxmlRels: excel.parseWorkbookRels(),
            workbookxml: excel.parseWorkbook(),
            sheetXmls: excel.parseWorksheetsDir(),
            templateSheetRel: excel.templateSheetRel()
        }).then(({sharedstrings, workbookxmlRels,workbookxml,sheetXmls,templateSheetRel})=>{
            this.relationship = new WorkBookRels(workbookxmlRels);
            this.workbookxml = new WorkBookXml(workbookxml);
            this.sheetXmls = new SheetXmls(sheetXmls);
            this.sharedstrings = new SharedStrings(sharedstrings, this.sheetXmls.templateSheetData());
            return this;
        });
    }

    /**
     * merge
     * @param {Object} bindingData {key1:value, key2:value, key3:value ~}
     * @param {Object} option
     * @return {Object} excel data. Blob if on browser. Node-buffer if on Node.js.
     */
    merge(bindingData, option = {type: config.buffer_type_output, compression: config.compression}){
        return Excel.instanceOf(this.excel)
            .merge(bindingData)
            .generate(option);
    }

    /**
     * bulkMergeMultiFile
     * @param {Array} bindingDataArray [{name:fileName, data:bindingData},,,,,]
     * @return {Object} excel data. Blob if on browser. Node-buffer if on Node.js.
     */
    bulkMergeMultiFile(bindingDataArray){
        return _.reduce(bindingDataArray, (excel, {name, data}) => {
            excel.file(name, this.merge(data, {type: config.buffer_type_jszip, compression: config.compression}));
            return excel;
        }, new Excel())
        .generate({type: config.buffer_type_output, compression: config.compression});
    }

    /**
     * bulkMergeMultiSheet
     * @param {Array} bindingDataArray [{name:sheetName, data:bindingData},,,,,]
     * @return {Object} excel data. Blob if on browser. Node-buffer if on Node.js.
     */
    bulkMergeMultiSheet(bindingDataArray){
        _.each(bindingDataArray, ({name,data})=>this.addSheetBindingData(name,data));
        return this.generate({type: config.buffer_type_output, compression: config.compression});
    }

    /**
     * generate
     * @param {Object} option
     * @return {Object} excel data. Blob if on browser. Node-buffer if on Node.js.
     * @private
     */
    generate(option){
        this.deleteTemplateSheet();
        return this.excel
            .setSharedStrings(this.sharedstrings.value())
            .setWorkbookRels(this.relationship.value())
            .setWorkbook(this.workbookxml.value())
            .setWorksheets(this.sheetXmls.value())
            .setWorksheetRels(this.sheetXmls.names())
            .generate(option);
    }

    /**
     * addSheetBindingData
     * @param {String} destSheetName
     * @param {Object} bindingData {key1:value, key2:value, key3:value ~}
     * @return {Object} this
     * @private
     */
    addSheetBindingData(destSheetName, bindingData){
        let nextId = this.relationship.nextRelationshipId();
        this.relationship.add(nextId);
        this.workbookxml.add(destSheetName, nextId);
        this.sharedstrings.addMergedStrings(bindingData);

        let sourceSheet = this.findSheetByName(this.workbookxml.firstSheetName()).value;
        let addedSheet = this.buildNewSheet(sourceSheet, bindingData);

        this.sheetXmls.add(nextId, addedSheet);

        return this;
    }

    /**
     * deleteTemplateSheet
     * @private
     */
    deleteTemplateSheet(){
        let sheetname = this.workbookxml.firstSheetName();
        let targetSheet = this.findSheetByName(sheetname);
        this.relationship.delete(targetSheet.path);
        this.workbookxml.delete(sheetname);

        _.each(this.sheetXmls.value(), ({name, data})=>{
            if((name === targetSheet.value.name)) {
                this.excel.removeWorksheet(targetSheet.value.name);
                this.excel.removeWorksheetRel(targetSheet.value.name);
            }
        });
        this.sheetXmls.delete(targetSheet.value.name);
    }

    /**
     * buildNewSheet
     * @param {Object} sourceSheet
     * @param {Object} bindingData {key1:value, key2:value, key3:value ~}
     * @return {SheetXmls}
     * @private
     */
    buildNewSheet(sourceSheet, bindingData){
        let addedSheet = _.deepCopy(sourceSheet);
        addedSheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected = '0';
        this.setCellIndexes(addedSheet, bindingData);
        return addedSheet;
    }

    //TODO このメソッドはsheetXmlsのメソッドにうつす予定
    /**
     * setCellIndexes
     * @param {Object} sheet
     * @param {Object} bindingData {key1:value, key2:value, key3:value ~}
     */
    setCellIndexes(sheet, bindingData) {
        let mergedStrings = this.sharedstrings.buildNewSharedStrings(bindingData);
        _.each(mergedStrings,(string)=>{
            _.each(string.usingCells, (cellAddress)=>{
                _.each(sheet.worksheet.sheetData[0].row,(row)=>{
                    _.each(row.c,(cell)=>{
                        if(cell['$'].r === cellAddress){
                            cell.v[0] = string.sharedIndex;
                        }
                    });
                });
            });
        });
    }

    /**
     * findSheetByName
     * @param {String} sheetname
     * @param {Object}
     */
    findSheetByName(sheetname){
        let sheetid = this.workbookxml.findSheetId(sheetname);
        if(!sheetid){
            return null;
        }
        let targetFilePath = this.relationship.findSheetPath(sheetid);
        let targetFileName = _.last(targetFilePath.split('/'));
        return {path: targetFilePath, value: this.sheetXmls.find(targetFileName)};
    }

    /**
     * variables
     * @param {Object}
     */
    variables(){
        return this.excel.variables();
    }
}

module.exports = ExcelMerge;
