/**
 * SheetHelper
 * Manage MS-Excel file. core business-logic class for js-merge-xlsx.
 * @author Satoshi Haga
 * @date 2015/10/03
 */
const Promise = require('bluebird');
const _ = require('underscore');
require('./underscore_mixin');

const Excel = require('./Excel');
const WorkBookXml = require('./WorkBookXml');
const WorkBookRels = require('./WorkBookRels');
const SheetXmls = require('./SheetXmls');
const SharedStrings = require('./SharedStrings');

const isNode = require('detect-node');
const config = {
    compression: "DEFLATE",
    buffer_type_output: (isNode ? 'nodebuffer' : 'blob'),
    buffer_type_jszip: (isNode ? 'nodebuffer' : 'arraybuffer')
};

class SheetHelper{

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

    simpleMerge(mergedData, option = {type: config.buffer_type_output, compression: config.compression}){
        return Excel.instanceOf(this.excel)
            .merge(mergedData)
            .generate(option);
    }

    bulkMergeMultiFile(mergedDataArray){
        return _.reduce(mergedDataArray, (excel, {name, data}) => {
            excel.file(name, this.simpleMerge(data, {type: config.buffer_type_jszip, compression: config.compression}));
            return excel;
        }, new Excel()).generate({type: config.buffer_type_output, compression: config.compression});
    }

    bulkMergeMultiSheet(mergedDataArray){
        _.each(mergedDataArray, ({name,data})=>this.addSheetBindingData(name,data));
        return this.generate({type: config.buffer_type_output, compression: config.compression});
    }

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

    addSheetBindingData(destSheetName, mergedData){
        let nextId = this.relationship.nextRelationshipId();
        this.relationship.add(nextId);
        this.workbookxml.add(destSheetName, nextId);
        this.sharedstrings.addMergedStrings(mergedData);

        let sourceSheet = this.findSheetByName(this.workbookxml.firstSheetName()).value;
        let addedSheet = this.buildNewSheet(sourceSheet, mergedData);

        this.sheetXmls.add(nextId, addedSheet);

        return this;
    }

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

    buildNewSheet(sourceSheet, mergedData){
        let addedSheet = _.deepCopy(sourceSheet);
        addedSheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected = '0';
        this.setCellIndexes(addedSheet, mergedData);
        return addedSheet;
    }

    //TODO このメソッドはsheetXmlsのメソッドにうつす予定
    setCellIndexes(sheet, mergedData) {
        let mergedStrings = this.sharedstrings.buildNewSharedStrings(mergedData);
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

    findSheetByName(sheetname){
        let sheetid = this.workbookxml.findSheetId(sheetname);
        if(!sheetid){
            return null;
        }
        let targetFilePath = this.relationship.findSheetPath(sheetid);
        let targetFileName = _.last(targetFilePath.split('/'));
        return {path: targetFilePath, value: this.sheetXmls.find(targetFileName)};
    }

}

module.exports = SheetHelper;
