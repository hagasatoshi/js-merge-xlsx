/**
 * SheetHelper
 * Manage MS-Excel file. core business-logic class for js-merge-xlsx.
 * @author Satoshi Haga
 * @date 2015/10/03
 */
var Mustache = require('mustache');
var Promise = require('bluebird');
var _ = require('underscore');
require('./underscore_mixin');
var JSZip = require('jszip');
var isNode = require('detect-node');
const outputBuffer = {type: (isNode?'nodebuffer':'blob'), compression:"DEFLATE"};
const jszipBuffer = {type: (isNode?'nodebuffer':'arraybuffer'), compression:"DEFLATE"};
var xml2js = require('xml2js');
var parseString = Promise.promisify(xml2js.parseString);
var builder = new xml2js.Builder();

const OPEN_XML_SCHEMA_DEFINITION = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';

class SheetHelper{

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
        //set member variable
        this.excel = excel;
        this.variables = _.variables(excel.file('xl/sharedStrings.xml').asText());
        this.commonStringsWithVariable = [];

        //some members are parsed in promise-chain because xml2js parses asynchronously
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

    /**
     * simpleMerge
     * @param {Object} bindData binding data
     * @return {Promise} Promise instance including MS-Excel data.
     */
    simpleMerge(bindData){

        //validation
        if(!bindData){
            throw new Error('simpleMerge() must has parameter');
        }

        return Promise.resolve().then(()=>this._simpleMerge(bindData, outputBuffer));
    }

    /**
     * bulkMergeMultiFile
     * @param {Array} bindDataArray including data{name: file's name, data: binding-data}
     * @return {Promise} Promise instance including MS-Excel data.
     */
    bulkMergeMultiFile(bindDataArray){

        //validation
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

    /**
     * addSheetBindingData
     * @param {String} dest_sheet_name name of new sheet
     * @param {Object} data binding data
     * @return {Object} this instance for chaining
     */
    addSheetBindingData(destSheetName, data){
        //validation
        if((!destSheetName) || !(data)) {
            throw new Error('addSheetBindingData() needs to have 2 paramter.');
        }
        //1.add relation of next sheet
        let nextId = this._availableSheetid();
        this.workbookxmlRels.Relationships.Relationship.push({ '$': { Id: nextId, Type: OPEN_XML_SCHEMA_DEFINITION, Target: `worksheets/sheet${nextId}.xml`}});
        this.workbookxml.workbook.sheets[0].sheet.push({ '$': { name: destSheetName, sheetId: nextId.replace('rId',''), 'r:id': nextId } });

        //2.add sheet file.
        let mergedStrings;
        if(this.sharedstrings){

            //prepare merged-strings
            mergedStrings = _.deepCopy(this.commonStringsWithVariable);
            _.each(mergedStrings,(e)=>e.t[0] = Mustache.render(_.stringValue(e.t), data));

            //add merged-string into sharedstrings
            let currentCount = this.sharedstrings.length;
            _.each(mergedStrings,(e,index)=>{
                e.sharedIndex = currentCount + index;
                this.sharedstrings.push(e);
            });
        }

        //build new sheet oject
        let sourceSheet = this._sheetByName(this.templateSheetName).value;
        let addedSheet = this._buildNewSheet(sourceSheet, mergedStrings);

        //update sheet name.
        addedSheet.name = `sheet${nextId}.xml`;

        //add this sheet into sheet_xmls
        this.sheetXmls.push(addedSheet);

        return this;
    }

    /**
     * hasSheet
     * @param {String} sheetname target sheet name
     * @return {boolean}
     */
    hasSheet(sheetname){
        return !!this._sheetByName(sheetname);
    }

    /**
     * focusOnFirstSheet
     * @return {Object} this instance for chaining
     */
    focusOnFirstSheet(){
        let targetSheetName = this._sheetByName(this._firstSheetName());
        _.each(this.sheet_xmls, (sheet)=>{
            if(!sheet.worksheet) return;
            sheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected = (sheet.name === targetSheetName.value.worksheet.name) ? '1' : '0';
        });
        return this;

    }

    /**
     * isFocused
     * @param {String} sheetname target sheet name
     * @return {boolean}
     */
    isFocused(sheetname){

        //validation
        if(!sheetname){
            throw new Error('isFocused() needs to have 1 paramter.');
        }
        if(!this.hasSheet(sheetname)){
            throw new Error(`Invalid sheet name '${sheetname}'.`);
        }

        let targetSheetName = this._sheetByName(sheetname);
        return (targetSheetName.value.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected === '1');
    }

    /**
     * deleteSheet
     * @param {String} sheetname target sheet name
     * @return {Object} this instance for chaining
     */
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

    /**
     * deleteTemplateSheet
     * @return {Object} this instance for chaining
     */
    deleteTemplateSheet(){
        return this.deleteSheet(this.templateSheetName);
    }

    /**
     * hasAsSharedString
     * @param {String} targetStr
     * @return {boolean}
     */
    hasAsSharedString(targetStr){
        return (this.excel.file('xl/sharedStrings.xml').asText().indexOf(targetStr) !== -1)
    }

    /**
     * generate
     * call JSZip#generate() binding current data
     * @param {Object} option option for JsZip#genereate()
     * @return {Promise} Promise instance inclusing Excel data.
     */
    generate(option){
        return parseString(this.excel.file('xl/sharedStrings.xml').asText())
        .then((sharedstringsObj)=> {

            if (this.sharedstrings) {
                sharedstringsObj.sst.si = _.deleteProperties(this.sharedstrings, ['sharedIndex', 'usingCells']);
                sharedstringsObj.sst['$'].uniqueCount = this.sharedstrings.length;
                sharedstringsObj.sst['$'].count = this._stringCount();

                this.excel.file('xl/sharedStrings.xml', _.decode(builder.buildObject(sharedstringsObj)))
            }

            //workbook.xml.rels
            this.excel.file("xl/_rels/workbook.xml.rels",_.decode(builder.buildObject(this.workbookxmlRels)));

            //workbook.xml
            this.excel.file("xl/workbook.xml", _.decode(builder.buildObject(this.workbookxml)));

            //sheetXmls
            _.each(this.sheetXmls, (sheet)=>{
                if(sheet.name){
                    var sheetObj = {};
                    sheetObj.worksheet = {};
                    _.extend(sheetObj.worksheet, sheet.worksheet);
                    this.excel.file(`xl/worksheets/${sheet.name}`, _.decode(builder.buildObject(sheetObj)));
                }
            });

            //sheetXmlsRels
            if(this.templateSheetRelsData.value && this.templateSheetRelsData.value.Relationships){
                let strTemplateSheetRels = _.decode(builder.buildObject({Relationships:this.templateSheetRelsData.value.Relationships}));
                _.each(this.sheetXmls, (sheet)=>{
                    if(sheet.name) this.excel.file(`xl/worksheets/_rels/${sheet.name}.rels`, strTemplateSheetRels);
                });
            }

            //call JSZip#generate()
            return this.excel.generate(option);
        })

    }

    /**
     * _simpleMerge
     * @param {Object} bindData binding data
     * @param {Object} option JsZip#generate() option.
     * @returns {Object} rendered MS-Excel data. data-format is determined by jszip_option
     * @private
     */
    _simpleMerge(bindData, option=outputBuffer){
        return new JSZip(this.excel.generate(jszipBuffer))
            .file('xl/sharedStrings.xml', Mustache.render(this.excel.file('xl/sharedStrings.xml').asText(), bindData))
            .generate(option);
    }

    /**
     * _parseCommonStringWithVariable
     * @return {Array} including common strings only having mustache-variable
     * @private
     */
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

    /**
     * _parseDirInExcel
     * @param {String} dir directory name in Zip file.
     * @return {Promise|Array} array including files parsed by xml2js
     * @private
     */
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

    /**
     * _buildNewSheet
     * @param {Object} sourceSheet
     * @param {Array} commonStringsWithVariable
     * @return {Object}
     * @private
     */
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

    /**
     * _availableSheetid
     * @return {String} id of next sheet
     * @private
     */
    _availableSheetid(){
        let maxRel = _.max(this.workbookxmlRels.Relationships.Relationship, (e)=> Number(e['$'].Id.replace('rId','')));
        let nextId = 'rId' + ('00' + (((maxRel['$'].Id.replace('rId','') >> 0))+1)).slice(-3);
        return nextId;
    }

    /**
     * _sheetByName
     * @param {String} sheetname target sheet name
     * @return {Object} sheet object
     * @private
     */
    _sheetByName(sheetname){
        let targetSheet = _.find(this.workbookxml.workbook.sheets[0].sheet, (e)=> (e['$'].name === sheetname));
        if(!targetSheet) return null;  //invalid sheet name

        let sheetid = targetSheet['$']['r:id'];
        let targetFilePath = _.max(this.workbookxmlRels.Relationships.Relationship, (e)=>(e['$'].Id === sheetid))['$'].Target;
        let targetFileName = _.last(targetFilePath.split('/'));
        return {path: targetFilePath, value: _.find(this.sheetXmls, (e)=>(e.name === targetFileName))};
    }

    /**
     * _sheetRelsByName
     * @param {String} sheetname target sheet name
     * @return {Object} sheet_rels object
     * @private
     */
    _sheetRelsByName(sheetname){
        let targetFilePath = this._sheetByName(sheetname).path;
        let targetName = `${_.last(targetFilePath.split('/'))}.rels`;
        return {name: targetName, value: _.find(this.sheetXmlsRels, e=>(e.name === targetName))};
    }

    /**
     * _templateSheetRels
     * @return {Object} sheet_rels object of template-sheet
     * @private
     */
    _templateSheetRels(){
        return this._sheetRelsByName(this.templateSheetName);
    }

    /**
     * _firstSheetName
     * @return {String} name of first-sheet of MS-Excel file
     * @private
     */
    _firstSheetName(){
        return this.workbookxml.workbook.sheets[0].sheet[0]['$'].name;
    }

    /**
     * _stringCount
     * @return {Number} count of string-cell
     * @private
     */
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

//Exports
module.exports = SheetHelper;
