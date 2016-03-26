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
const JSZip = require('jszip');
const isNode = require('detect-node');
const outputBuffer = {type: (isNode?'nodebuffer':'blob'), compression:"DEFLATE"};
const jszipBuffer = {type: (isNode?'nodebuffer':'arraybuffer'), compression:"DEFLATE"};
const xml2js = require('xml2js');
const parseString = Promise.promisify(xml2js.parseString);
const builder = new xml2js.Builder();

const OPEN_XML_SCHEMA_DEFINITION = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';

class SheetHelper{

    load(excel){
        if(!(excel instanceof JSZip)){
            return Promise.reject('First parameter must be JSZip instance including MS-Excel data');
        }
        this.excel = excel;
        this.variables = _.variables(excel.file('xl/sharedStrings.xml').asText());
        this.commonStringsWithVariable = [];

        return Promise.props({
            sharedstringsObj: parseString(excel.file('xl/sharedStrings.xml').asText()),
            workbookxmlRels: parseString(this.excel.file('xl/_rels/workbook.xml.rels').asText()),
            workbookxml: parseString(this.excel.file('xl/workbook.xml').asText()),
            sheetXmls: this._parseDirInExcel('xl/worksheets'),
            sheetXmlsRels: this._parseDirInExcel('xl/worksheets/_rels')
        }).then(({sharedstringsObj, workbookxmlRels,workbookxml,sheetXmls,sheetXmlsRels})=>{
            this.sharedstrings = sharedstringsObj.sst.si;
            this.workbookxmlRels = workbookxmlRels;
            this.workbookxml = workbookxml;
            this.sheetXmls = sheetXmls;
            this.sheetXmlsRels = sheetXmlsRels;
            this.templateSheetName = this.workbookxml.workbook.sheets[0].sheet[0]['$'].name;
            this.templateSheetData = _.find(sheetXmls,(e)=>(e.name.indexOf('.rels') === -1)).worksheet.sheetData[0].row;
            this.templateSheetRelsData = _.deepCopy(this._templateSheetRels());
            this.commonStringsWithVariable = this._parseCommonStringWithVariable();
            //return this for chaining
            return this;
        });
    }

    simpleMerge(bindData){
        if(!bindData){
            throw new Error('simpleMerge() must has parameter');
        }

        return Promise.resolve().then(()=>this._simpleMerge(bindData, outputBuffer));
    }

    bulkMergeMultiFile(bindDataArray){
        if(!_.isArray(bindDataArray)){
            throw new Error('bulkMergeMultiFile() has only array object');
        }
        if(_.find(bindDataArray,(e)=>!(e.name && e.data))){
            throw new Error('bulkMergeMultiFile() is called with invalid parameter');
        }

        var allExcels = new JSZip();
        _.each(bindDataArray, ({name,data})=>allExcels.file(name, this._simpleMerge(data, jszipBuffer)));
        return Promise.resolve().then(()=> allExcels.generate(outputBuffer));
    }

    addSheetBindingData(destSheetName, data){
        if((!destSheetName) || !(data)) {
            throw new Error('addSheetBindingData() needs to have 2 paramter.');
        }
        let nextId = this._availableSheetid();
        this.workbookxmlRels.Relationships.Relationship.push({ '$': { Id: nextId, Type: OPEN_XML_SCHEMA_DEFINITION, Target: `worksheets/sheet${nextId}.xml`}});
        this.workbookxml.workbook.sheets[0].sheet.push({ '$': { name: destSheetName, sheetId: nextId.replace('rId',''), 'r:id': nextId } });

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

        let sourceSheet = this._sheetByName(this.templateSheetName).value;
        let addedSheet = this._buildNewSheet(sourceSheet, mergedStrings);

        addedSheet.name = `sheet${nextId}.xml`;

        this.sheetXmls.push(addedSheet);

        return this;
    }

    hasSheet(sheetname){
        return !!this._sheetByName(sheetname);
    }

    focusOnFirstSheet(){
        let targetSheetName = this._sheetByName(this._firstSheetName());
        _.each(this.sheet_xmls, (sheet)=>{
            if(!sheet.worksheet) return;
            sheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected = (sheet.name === targetSheetName.value.worksheet.name) ? '1' : '0';
        });
        return this;

    }

    isFocused(sheetname){
        if(!sheetname){
            throw new Error('isFocused() needs to have 1 paramter.');
        }
        if(!this.hasSheet(sheetname)){
            throw new Error(`Invalid sheet name '${sheetname}'.`);
        }

        let targetSheetName = this._sheetByName(sheetname);
        return (targetSheetName.value.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected === '1');
    }

    deleteSheet(sheetname){
        if(!sheetname){
            throw new Error('deleteSheet() needs to have 1 paramter.');
        }
        let targetSheet = this._sheetByName(sheetname);
        if(!targetSheet){
            throw new Error(`Invalid sheet name '${sheetname}'.`);
        }
        _.each(this.workbookxmlRels.Relationships.Relationship, (sheet,index)=>{
            if(sheet && (sheet['$'].Target === targetSheet.path)) this.workbookxmlRels.Relationships.Relationship.splice(index,1);
        });
        _.each(this.workbookxml.workbook.sheets[0].sheet, (sheet,index)=>{
            if(sheet && (sheet['$'].name === sheetname))this.workbookxml.workbook.sheets[0].sheet.splice(index,1);
        });
        _.each(this.sheetXmls, (sheetXml,index)=>{
            if(sheetXml && (sheetXml.name === targetSheet.value.name)) {
                this.sheetXmls.splice(index,1);
                this.excel.remove(`xl/worksheets/${targetSheet.value.name}`);
                this.excel.remove(`xl/worksheets/_rels/${targetSheet.value.name}.rels`);
            }
        });
        return this;
    }

    deleteTemplateSheet(){
        return this.deleteSheet(this.templateSheetName);
    }

    hasAsSharedString(targetStr){
        return (this.excel.file('xl/sharedStrings.xml').asText().indexOf(targetStr) !== -1)
    }

    templateVariables(){
        return this.variables;
    }

    generate(option){
        return parseString(this.excel.file('xl/sharedStrings.xml').asText())
        .then((sharedstringsObj)=> {

            if (this.sharedstrings) {
                sharedstringsObj.sst.si = _.deleteProperties(this.sharedstrings, ['sharedIndex', 'usingCells']);
                sharedstringsObj.sst['$'].uniqueCount = this.sharedstrings.length;
                sharedstringsObj.sst['$'].count = this._stringCount();

                this.excel.file('xl/sharedStrings.xml', _.decode(builder.buildObject(sharedstringsObj)))
            }
            this.excel.file("xl/_rels/workbook.xml.rels",_.decode(builder.buildObject(this.workbookxmlRels)));
            this.excel.file("xl/workbook.xml", _.decode(builder.buildObject(this.workbookxml)));
            _.each(this.sheetXmls, (sheet)=>{
                if(sheet.name){
                    var sheetObj = {};
                    sheetObj.worksheet = {};
                    _.extend(sheetObj.worksheet, sheet.worksheet);
                    this.excel.file(`xl/worksheets/${sheet.name}`, _.decode(builder.buildObject(sheetObj)));
                }
            });
            if(this.templateSheetRelsData.value && this.templateSheetRelsData.value.Relationships){
                let strTemplateSheetRels = _.decode(builder.buildObject({Relationships:this.templateSheetRelsData.value.Relationships}));
                _.each(this.sheetXmls, (sheet)=>{
                    if(sheet.name) this.excel.file(`xl/worksheets/_rels/${sheet.name}.rels`, strTemplateSheetRels);
                });
            }

            return this.excel.generate(option);
        })

    }

    _simpleMerge(bindData, option=outputBuffer){
        return new JSZip(this.excel.generate(jszipBuffer))
            .file('xl/sharedStrings.xml', Mustache.render(this.excel.file('xl/sharedStrings.xml').asText(), bindData))
            .generate(option);
    }

    _parseCommonStringWithVariable(){
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

    _parseDirInExcel(dir){
        let files = this.excel.folder(dir).file(/.xml/);
        let fileXmls = [];
        return files.reduce(
            (promise, file)=>
                promise.then((prior_file)=>
                    Promise.resolve()
                        .then(()=>parseString(this.excel.file(file.name).asText()))
                        .then((file_xml)=>{
                            file_xml.name = _.last(file.name.split('/'));
                            fileXmls.push(file_xml);
                            return fileXmls;
                        })
                )
            ,
            Promise.resolve()
        );
    }

    _buildNewSheet(sourceSheet, commonStringsWithVariable){
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

    _availableSheetid(){
        let maxRel = _.max(this.workbookxmlRels.Relationships.Relationship, (e)=> Number(e['$'].Id.replace('rId','')));
        let nextId = 'rId' + ('00' + (((maxRel['$'].Id.replace('rId','') >> 0))+1)).slice(-3);
        return nextId;
    }

    _sheetByName(sheetname){
        let targetSheet = _.find(this.workbookxml.workbook.sheets[0].sheet, (e)=> (e['$'].name === sheetname));
        if(!targetSheet) return null;  //invalid sheet name

        let sheetid = targetSheet['$']['r:id'];
        let targetFilePath = _.max(this.workbookxmlRels.Relationships.Relationship, (e)=>(e['$'].Id === sheetid))['$'].Target;
        let targetFileName = _.last(targetFilePath.split('/'));
        return {path: targetFilePath, value: _.find(this.sheetXmls, (e)=>(e.name === targetFileName))};
    }

    _sheetRelsByName(sheetname){
        let targetFilePath = this._sheetByName(sheetname).path;
        let targetName = `${_.last(targetFilePath.split('/'))}.rels`;
        return {name: targetName, value: _.find(this.sheetXmlsRels, e=>(e.name === targetName))};
    }

    _templateSheetRels(){
        return this._sheetRelsByName(this.templateSheetName);
    }

    _firstSheetName(){
        return this.workbookxml.workbook.sheets[0].sheet[0]['$'].name;
    }

    _stringCount(){
        let stringCount = 0;
        _.each(this.sheetXmls, (sheet)=>{
            if(sheet.worksheet){
                _.each(sheet.worksheet.sheetData[0].row, (row)=>{
                    _.each(row.c, (cell)=>{
                        if(cell['$'].t){
                            stringCount++;
                        }
                    });
                });
            }
        });
        return stringCount;
    }
}

module.exports = SheetHelper;
