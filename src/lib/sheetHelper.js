/**
 * SheetHelper
 * Manage MS-Excel file. core business-logic class for js-merge-xlsx.
 * @author Satoshi Haga
 * @date 2015/10/03
 */
const Mustache = require('mustache');
const Promise = require('bluebird');
const _ = require('underscore');
require('./underscore_mixin');
const Excel = require('./Excel');
const WorkBookXml = require('./WorkBookXml');
const WorkBookRels = require('./WorkBookRels');
const SheetXmls = require('./SheetXmls');
const SharedStrings = require('./SharedStrings');
const isNode = require('detect-node');
const outputBuffer = {type: (isNode?'nodebuffer':'blob'), compression:"DEFLATE"};
const jszipBuffer = {type: (isNode?'nodebuffer':'arraybuffer'), compression:"DEFLATE"};

class SheetHelper{

    load(excel){
        this.excel = excel;
        this.commonStringsWithVariable = [];
        return Promise.props({
            sharedstringsObj: excel.parseSharedStrings(),
            workbookxmlRels: excel.parseWorkbookRels(),
            workbookxml: excel.parseWorkbook(),
            sheetXmls: excel.parseWorksheetsDir(),
            templateSheetRel: excel.templateSheetRel()
        }).then(({sharedstringsObj, workbookxmlRels,workbookxml,sheetXmls,templateSheetRel})=>{
            this.sharedstrings = new SharedStrings(sharedstringsObj);
            this.relationship = new WorkBookRels(workbookxmlRels);
            this.workbookxml = new WorkBookXml(workbookxml);
            this.sheetXmls = new SheetXmls(sheetXmls);
            this.templateSheetRel = templateSheetRel;
            this.templateSheetData = _.find(sheetXmls,(e)=>(e.name.indexOf('.rels') === -1)).worksheet.sheetData[0].row;
            this.commonStringsWithVariable = this.parseCommonStringWithVariable();
            //return this for chaining
            return this;
        });
    }

    simpleMerge(mergedData, option=outputBuffer){
        return Excel.instanceOf(this.excel)
            .merge(mergedData)
            .generate(option);
    }

    bulkMergeMultiFile(mergedDataArray){
        return _.reduce(mergedDataArray, (excel, {name, data}) => {
            excel.file(name, this.simpleMerge(data, jszipBuffer));
            return excel;
        }, new Excel()).generate(outputBuffer);
    }

    bulkMergeMultiSheet(mergedDataArray){
        _.each(mergedDataArray, ({name,data})=>this.addSheetBindingData(name,data));
        return this.generate(outputBuffer);
    }


    addSheetBindingData(destSheetName, data){
        let nextId = this.relationship.nextRelationshipId();
        this.relationship.add(nextId);
        this.workbookxml.add(destSheetName, nextId);

        let mergedStrings;
        if(this.sharedstrings.hasString()){

            mergedStrings = _.deepCopy(this.commonStringsWithVariable);
            _.each(mergedStrings,(e)=>e.t[0] = Mustache.render(_.stringValue(e.t), data));

            this.sharedstrings.add(mergedStrings);
        }

        let sourceSheet = this.sheetByName(this.workbookxml.firstSheetName()).value;
        let addedSheet = this.buildNewSheet(sourceSheet, mergedStrings);

        addedSheet.name = `sheet${nextId}.xml`;

        this.sheetXmls.add(addedSheet);

        return this;
    }

    hasSheet(sheetname){
        return !!this.sheetByName(sheetname);
    }

    isFocused(sheetname){
        let targetSheetName = this.sheetByName(sheetname);
        return (targetSheetName.value.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected === '1');
    }

    deleteTemplateSheet(){
        let sheetname = this.workbookxml.firstSheetName();
        let targetSheet = this.sheetByName(sheetname);
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

    generate(option){
        this.deleteTemplateSheet();
        return this.excel
        .setSharedStrings(this.sharedstrings.value())
        .setWorkbookRels(this.relationship.value())
        .setWorkbook(this.workbookxml.value())
        .setWorksheets(this.sheetXmls.value())
        .setWorksheetRels(this.sheetXmls.names(), this.templateSheetRel)
        .generate(option);
    }


    parseCommonStringWithVariable(){
        let commonStringsWithVariable = this.sharedstrings.filterWithVariable();

        _.each(commonStringsWithVariable, (commonStringWithVariable)=>{
            commonStringWithVariable.usingCells = [];
            _.each(this.templateSheetData,(row)=>{
                _.each(row.c,(cell)=>{
                    if(cell['$'].t === 's'){
                        if(commonStringWithVariable.sharedIndex === (cell.v[0] >> 0)){
                            commonStringWithVariable.usingCells.push(cell['$'].r);
                        }
                    }
                });
            });
        });

        return commonStringsWithVariable;
    }

    buildNewSheet(sourceSheet, commonStringsWithVariable){
        let addedSheet = _.deepCopy(sourceSheet);
        addedSheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected = '0';
        if(!commonStringsWithVariable) return addedSheet;

        _.each(commonStringsWithVariable,(e,index)=>{
            _.each(e.usingCells, (cellAddress)=>{
                _.each(addedSheet.worksheet.sheetData[0].row,(row)=>{
                    _.each(row.c,(cell)=>{
                        if(cell['$'].r === cellAddress){
                            cell.v[0] = e.sharedIndex;
                        }
                    });
                });
            });
        });
        return addedSheet;
    }

    sheetByName(sheetname){
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
