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
    compression:        'DEFLATE',
    buffer_type_output: (isNode ? 'nodebuffer' : 'blob'),
    buffer_type_jszip:  (isNode ? 'nodebuffer' : 'arraybuffer')
};

class ExcelMerge {

    /**
     * load
     * @param {Excel} excel
     * @return {Promise} this
     */
    load(excel) {
        this.excel = excel;
        return excel.setTemplateSheetRel()
        .then(() => {
            return Promise.props({
                sharedstrings:   excel.parseSharedStrings(),
                workbookxmlRels: excel.parseWorkbookRels(),
                workbookxml:     excel.parseWorkbook(),
                sheetXmls:       excel.parseWorksheetsDir()
            })
        }).then(({sharedstrings, workbookxmlRels, workbookxml, sheetXmls}) => {
            this.relationship = new WorkBookRels(workbookxmlRels);
            this.workbookxml = new WorkBookXml(workbookxml);
            this.sheetXmls = new SheetXmls(sheetXmls);
            this.templateSheetModel = this.sheetXmls.getTemplateSheetModel();
            this.sharedstrings = new SharedStrings(
                sharedstrings, this.sheetXmls.templateSheetData()
            );
            return this;
        });
    }

    /**
     * merge
     * @param {Object} bindingData {key1:value, key2:value, key3:value ~}
     * @param {Object} option
     * @return {Object} excel data. Blob if on browser. Node-buffer if on Node.js.
     */
    merge(
        bindingData, option = {type: config.buffer_type_output, compression: config.compression}
    ) {
        return Excel.instanceOf(this.excel)
            .merge(bindingData)
            .generate(option);
    }

    /**
     * bulkMergeMultiFile
     * @param {Array} bindingDataArray [{name:fileName, data:bindingData},,,,,]
     * @return {Object} excel data. Blob if on browser. Node-buffer if on Node.js.
     */
    bulkMergeMultiFile(bindingDataArray) {
        return _.reduce(bindingDataArray, (excel, {name, data}) => {
            excel.file(
                name,
                this.merge(
                    data, {type: config.buffer_type_jszip, compression: config.compression}
                )
            );
            return excel;
        }, new Excel())
        .generate({type: config.buffer_type_output, compression: config.compression});
    }

    /**
     * bulkMergeMultiSheet
     * @param {Array} bindingDataArray [{name:sheetName, data:bindingData},,,,,]
     * @return {Object} excel data. Blob if on browser. Node-buffer if on Node.js.
     */
    bulkMergeMultiSheet(bindingDataArray) {
        return _.reduce(
            bindingDataArray,
            (thisObj, {name, data}) => thisObj.addMergedSheet(name, data),
            this
        ).deleteTemplateSheet()
        .generate({type: config.buffer_type_output, compression: config.compression});
    }

    /**
     * generate
     * @param {Object} option
     * @return {Object} excel data. Blob if on browser. Node-buffer if on Node.js.
     * @private
     */
    generate(option) {
        return this.excel
            .setSharedStrings(this.sharedstrings.value())
            .setWorkbookRels(this.relationship.value())
            .setWorkbook(this.workbookxml.value())
            .setWorksheets(this.sheetXmls.value())
            .setWorksheetRels(this.sheetXmls.names())
            .generate(option);
    }

    /**
     * addMergedSheet
     * @param {String} newSheetName
     * @param {Object} mergeData {key1:value, key2:value, key3:value ~}
     * @return {Object} this
     * @private
     */
    addMergedSheet(newSheetName, mergeData) {
        let nextId = this.relationship.nextRelationshipId();
        this.relationship.add(nextId);
        this.workbookxml.add(newSheetName, nextId);
        this.sheetXmls.add(
            `sheet${nextId}.xml`,
            this.templateSheetModel.cloneWithMergedString(
                this.sharedstrings.addMergedStrings(mergeData)
            )
        );
        return this;
    }

    /**
     * deleteTemplateSheet
     * @private
     */
    deleteTemplateSheet() {
        let sheetname = this.workbookxml.firstSheetName();
        let targetSheet = this.findSheetByName(sheetname);
        this.relationship.delete(targetSheet.path);
        this.workbookxml.delete(sheetname);

        this.sheetXmls.delete(targetSheet.value.name);
        this.excel.removeWorksheet(targetSheet.value.name);
        return this;
    }

    /**
     * findSheetByName
     * @param {String} sheetname
     * @param {Object}
     */
    findSheetByName(sheetname) {
        let sheetid = this.workbookxml.findSheetId(sheetname);
        if(!sheetid) {
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
    variables() {
        return this.excel.variables();
    }
}

module.exports = ExcelMerge;
