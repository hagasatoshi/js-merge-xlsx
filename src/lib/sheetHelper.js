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
const isNode = require('detect-node');
const outputBuffer = {type: (isNode?'nodebuffer':'blob'), compression:"DEFLATE"};
const jszipBuffer = {type: (isNode?'nodebuffer':'arraybuffer'), compression:"DEFLATE"};

const OPEN_XML_SCHEMA_DEFINITION = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';

class SheetHelper{

    load(excel){
        if(!(excel instanceof Excel)){
            return Promise.reject('First parameter must be Excel instance including MS-Excel data');
        }
        this.excel = excel;
        this.variables = _.variables(excel.sharedStrings());
        this.commonStringsWithVariable = [];

        return Promise.props({
            sharedstringsObj: excel.parseSharedStrings(),
            workbookxmlRels: excel.parseWorkbookRels(),
            workbookxml: excel.parseWorkbook(),
            sheetXmls: excel.parseWorksheetsDir(),
            sheetXmlsRels: excel.parseWorksheetRelsDir()
        }).then(({sharedstringsObj, workbookxmlRels,workbookxml,sheetXmls,sheetXmlsRels})=>{
            this.sharedstrings = sharedstringsObj.sst.si;
            this.relationship = new WorkBookRels(workbookxmlRels);
            this.workbookxml = new WorkBookXml(workbookxml);
            this.sheetXmls = new SheetXmls(sheetXmls);
            this.sheetXmlsRels = sheetXmlsRels;
            this.templateSheetData = _.find(sheetXmls,(e)=>(e.name.indexOf('.rels') === -1)).worksheet.sheetData[0].row;
            this.templateSheetRelsData = _.deepCopy(this.templateSheetRels());
            this.commonStringsWithVariable = this.parseCommonStringWithVariable();
            //return this for chaining
            return this;
        });
    }

    simpleMerge(bindData){
        if(!bindData){
            throw new Error('simpleMerge() must has parameter');
        }

        return Promise.resolve().then(()=>this.simpleMerge(bindData, outputBuffer));
    }

    bulkMergeMultiFile(bindDataArray){
        if(!_.isArray(bindDataArray)){
            throw new Error('bulkMergeMultiFile() has only array object');
        }
        if(_.find(bindDataArray,(e)=>!(e.name && e.data))){
            throw new Error('bulkMergeMultiFile() is called with invalid parameter');
        }

        var allExcels = new Excel();
        _.each(bindDataArray, ({name,data})=>allExcels.file(name, this.simpleMerge(data, jszipBuffer)));
        return Promise.resolve().then(()=> allExcels.generate(outputBuffer));
    }

    addSheetBindingData(destSheetName, data){
        if((!destSheetName) || !(data)) {
            throw new Error('addSheetBindingData() needs to have 2 paramter.');
        }
        let nextId = this.relationship.nextRelationshipId();
        this.relationship.add(nextId);
        this.workbookxml.add(destSheetName, nextId);

        let mergedStrings;
        if(this.sharedstrings){

            mergedStrings = _.deepCopy(this.commonStringsWithVariable);
            _.each(mergedStrings,(e)=>e.t[0] = Mustache.render(_.stringValue(e.t), data));

            let currentCount = this.sharedstrings.length;
            _.each(mergedStrings,(e,index)=>{
                e.sharedIndex = currentCount + index;
                this.sharedstrings.push(e);
            });
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
        if(!sheetname){
            throw new Error('isFocused() needs to have 1 paramter.');
        }
        if(!this.hasSheet(sheetname)){
            throw new Error(`Invalid sheet name '${sheetname}'.`);
        }

        let targetSheetName = this.sheetByName(sheetname);
        return (targetSheetName.value.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected === '1');
    }

    deleteSheet(sheetname){
        if(!sheetname){
            throw new Error('deleteSheet() needs to have 1 paramter.');
        }
        let targetSheet = this.sheetByName(sheetname);
        if(!targetSheet){
            throw new Error(`Invalid sheet name '${sheetname}'.`);
        }
        this.relationship.delete(targetSheet.path);
        this.workbookxml.delete(sheetname);

        _.each(this.sheetXmls.value(), ({name, data})=>{
            if((name === targetSheet.value.name)) {
                this.excel.removeWorksheet(targetSheet.value.name);
                this.excel.removeWorksheetRel(targetSheet.value.name);
            }
        });
        this.sheetXmls.delete(targetSheet.value.name);
        return this;
    }

    deleteTemplateSheet(){
        return this.deleteSheet(this.workbookxml.firstSheetName());
    }

    templateVariables(){
        return this.variables;
    }

    generate(option){
        return this.excel.parseSharedStrings()
        .then((sharedstringsObj)=> {

            if (this.sharedstrings) {
                sharedstringsObj.sst.si = _.deleteProperties(this.sharedstrings, ['sharedIndex', 'usingCells']);
                sharedstringsObj.sst['$'].uniqueCount = this.sharedstrings.length;
                sharedstringsObj.sst['$'].count = this.sheetXmls.stringCount();

                this.excel.setSharedStrings(sharedstringsObj);
            }
            this.excel.setWorkbookRels(this.relationship.value());
            this.excel.setWorkbook(this.workbookxml.value());

            this.excel.setWorksheets(this.sheetXmls.value());

            if(this.templateSheetRelsData.value && this.templateSheetRelsData.value.Relationships){
                _.each(this.sheetXmls.value(), ({name, data})=>{
                    if(name){
                        this.excel.setWorksheetRel(name, { Relationships: this.templateSheetRelsData.value.Relationships });
                    }
                });
            }

            return this.excel.generate(option);
        })

    }

    simpleMerge(bindData, option=outputBuffer){
        return new Excel(this.excel.generate(jszipBuffer))
            .file('xl/sharedStrings.xml', Mustache.render(this.excel.sharedStrings(), bindData))
            .generate(option);
    }

    parseCommonStringWithVariable(){
        let commonStringsWithVariable = [];
        _.each(this.sharedstrings,(stringObj, index)=>{
            if(_.stringValue(stringObj.t) && _.hasVariable(_.stringValue(stringObj.t))){
                stringObj.sharedIndex = index;
                commonStringsWithVariable.push(stringObj);
            }
        });
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

    sheetRelsByName(sheetname){
        let targetFilePath = this.sheetByName(sheetname).path;
        let targetName = `${_.last(targetFilePath.split('/'))}.rels`;
        return {name: targetName, value: _.find(this.sheetXmlsRels, e=>(e.name === targetName))};
    }

    templateSheetRels(){
        return this.sheetRelsByName(this.workbookxml.firstSheetName());
    }

}

module.exports = SheetHelper;
