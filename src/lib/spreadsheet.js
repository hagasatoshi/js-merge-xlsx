/**
 * * SpreadSheet
 * * Manage MS-Excel file. core business-logic class for js-merge-xlsx.
 * * @author Satoshi Haga
 * * @date 2015/10/03
 **/
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

class SpreadSheet{

    /**
     * * member variables
     * * excel {Object} JSZip instance including template excel file
     * * variables {Array} including mustache-variables defined in sharedstrings.xml
     * * sharedstrings {Array} includings common strings defined in sharedstrings.xml
     * * sharedstrings_obj {Object} whole sharedstrings object
     * * commonStringsWithVariable {Array} including common strings only having mustache variables
     * * sheetXmls {Array} including objects parsed from  'xl/worksheets/*.xml'
     * * sheetXmlsRels {Array} including objects pared from 'xl/worksheets/_rels/*.xml.rels'
     * * templateSheetData {Object} object parsed from 'xl/worksheets/*.xml'. this is used as template-file
     * * templateSheetName {String} sheet-name of template-file
     * * workbookxmlRels {Object} parsed from 'xl/_rels/workbook.xml.rels'
     * * workbookxml {Object} parsed from 'xl/workbook.xml'
     * */


    /**
     * * load
     * * @param {Object} excel JsZip object including MS-Excel file
     * * @return {Promise|Object} Promise instance including this
     **/
    load(excel){
        //validation
        if(!(excel instanceof JSZip)) return Promise.reject('First parameter must be JSZip instance including MS-Excel data');
        //set member variable
        this.excel = excel;
        this.variables = _(excel.file('xl/sharedStrings.xml').asText()).variables();
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
            this.templateSheetRelsData = _(this._templateSheetRels()).deepCopy();
            this.commonStringsWithVariable = this._parseCommonStringWithVariable();
            //return this for chaining
            return this;
        });
    }

    /**
     * * simpleRender
     * * @param {Object} bind_data binding data
     * * @returns {Promise|Object} rendered MS-Excel data. data-format is determined by jszip_option
     **/
    simpleRender(bindData){

        //validation
        if(!bindData) return Promise.reject('simpleRender() must has parameter');

        return Promise.resolve().then(()=>this._simpleRender(bindData, outputBuffer));
    }

    /**
     * * bulkRenderMultiFile
     * * @param {Array} bindDataArray including data{name: file's name, data: binding-data}
     * * @returns {Promise|Object} rendered MS-Excel data.
     **/
    bulkRenderMultiFile(bindDataArray){

        //validation
        if(!_.isArray(bindDataArray)) return Promise.reject('bulkRenderMultiFile() has only array object');
        if(_.find(bindDataArray,(e)=>!(e.name && e.data))) return Promise.reject('bulkRenderMultiFile() is called with invalid parameter');

        var allExcels = new JSZip();
        _.each(bindDataArray, ({name,data})=>allExcels.file(name, this._simpleRender(data, jszipBuffer)));
        return Promise.resolve().then(()=> allExcels.generate(outputBuffer));
    }

    /**
     * * addSheetBindingData
     * * @param {String} dest_sheet_name name of new sheet
     * * @param {Object} data binding data
     * * @return {Object} this instance for chaining
     **/
    addSheetBindingData(destSheetName, data){
        //validation
        if((!destSheetName) || !(data)) return Promise.reject('addSheetBindingData() needs to have 2 paramter.');
        //1.add relation of next sheet
        let nextId = this._availableSheetid();
        this.workbookxmlRels.Relationships.Relationship.push({ '$': { Id: nextId, Type: OPEN_XML_SCHEMA_DEFINITION, Target: `worksheets/sheet${nextId}.xml`}});
        this.workbookxml.workbook.sheets[0].sheet.push({ '$': { name: destSheetName, sheetId: nextId.replace('rId',''), 'r:id': nextId } });

        //2.add sheet file.
        //2-1.prepare rendered-strings
        let renderedStrings = _(this.commonStringsWithVariable).deepCopy();
        _.each(renderedStrings,(e)=>e.t[0] = Mustache.render(_(e.t).stringValue(), data));

        //2-2.add rendered-string into sharedstrings
        let currentCount = this.sharedstrings.length;
        _.each(renderedStrings,(e,index)=>{
            e.sharedIndex = currentCount + index;
            this.sharedstrings.push(e);
        });

        //2-4.build new sheet oject
        let sourceSheet = this._sheetByName(this.templateSheetName).value;
        let addedSheet = this._buildNewSheet(sourceSheet, renderedStrings);

        //2-5.update sheet name.
        addedSheet.name = `sheet${nextId}.xml`;

        //2-6.add this sheet into sheet_xmls
        this.sheetXmls.push(addedSheet);

        return this;
    }

    /**
     * * activateSheet
     * * @param {String} sheetname target sheet name
     * * @return {Object} this instance for chaining
     **/
    activateSheet(sheetname){

        //validation
        if(!sheetname) return Promise.reject('activateSheet() needs to have 1 paramter.');

        let targetSheetName = this._sheetByName(sheetname);
        if(!targetSheetName) return Promise.reject(`Invalid sheet name '${sheetname}'.`);

        _.each(this.sheet_xmls, (sheet)=>{
            if(!sheet.worksheet) return;
            sheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected = (sheet.name === targetSheetName.value.worksheet.name) ? '1' : '0';
        });
        return this;
    }

    /**
     * * forcusOnFirstSheet
     * * @return {Object} this instance for chaining
     **/
    forcusOnFirstSheet(){
        return this.activateSheet(this._firstSheetName());
    }

    /**
     * * deleteSheet
     * * @param {String} sheetname target sheet name
     * * @return {Object} this instance for chaining
     **/
    deleteSheet(sheetname){
        if(!sheetname) return Promise.reject('deleteSheet() needs to have 1 paramter.');
        let targetSheet = this._sheetByName(sheetname);
        if(!targetSheet) return Promise.reject(`Invalid sheet name '${sheetname}'.`);
        _.each(this.workbookxmlRels.Relationships.Relationship, (sheet,index)=>{
            if(sheet && (sheet['$'].Target === targetSheet.path)) this.workbookxmlRels.Relationships.Relationship.splice(index,1);
        });
        _.each(this.workbookxml.workbook.sheets[0].sheet, (sheet,index)=>{
            if(sheet && (sheet['$'].name === sheetname))this.workbookxml.workbook.sheets[0].sheet.splice(index,1);
        });
        _.each(this.sheetXmls, (sheetXml,index)=>{
            if(sheetXml && (sheetXml.name === targetSheet.value.name)) this.sheetXmls.splice(index,1);
        });
        return this;
    }

    /**
     * * deleteTemplateSheet
     * * @return {Object} this instance for chaining
     **/
    deleteTemplateSheet(){
        return this.deleteSheet(this.templateSheetName);
    }

    /**
     * * hasAsSharedString
     * * @param {String} targetStr
     * * @return {boolean}
     **/
    hasAsSharedString(targetStr){
        return (this.excel.file('xl/sharedStrings.xml').asText().indexOf(targetStr) !== -1)
    }

    /**
     * * generate
     * * call JSZip#generate() binding current data
     * * @param {Object} option option for JsZip#genereate()
     * * @return {Object} Excel data. format is determinated by parameter
     **/
    generate(option){
        parseString(this.excel.file('xl/sharedStrings.xml').asText())
        .then((sharedstringsObj)=>{

            //sharedstring
            sharedstringsObj.sst.si = this._cleanSharedStrings();
            sharedstringsObj.sst['$'].count = sharedstringsObj.sst['$'].uniqueCount = this.sharedstrings.length;
            this.excel
                .file('xl/sharedStrings.xml', builder.buildObject(sharedstringsObj))
                .file("xl/_rels/workbook.xml.rels",builder.buildObject(this.workbookxmlRels))
                .file("xl/workbook.xml",builder.buildObject(this.workbookxml));

            //sheetXmls
            _.each(this.sheetXmls, (sheet)=>{
                if(sheet.name){
                    var sheetObj = {};
                    sheetObj.worksheet = {};
                    _.extend(sheetObj.worksheet, sheet.worksheet);
                    this.excel.file(`xl/worksheets/${sheet.name}`, builder.buildObject(sheetObj));
                }
            });

            //sheetXmlsRels
            let strTemplateSheetRels = builder.buildObject(this.templateSheetRelsData);
            _.each(this.sheetXmls, (sheet)=>{
                if(sheet.name) this.excel.file(`xl/worksheets/_rels/${sheet.name}.rels`, strTemplateSheetRels);
            });

            //call JSZip#generate()
            return this.excel.generate(option);
        })

    }


    /**
     * * _simpleRender
     * * @param {Object} bindData binding data
     * * @param {Object} option JsZip#generate() option.
     * * @returns {Object} rendered MS-Excel data. data-format is determined by jszip_option
     * * @private
     **/
    _simpleRender(bindData, option=outputBuffer){
        return new JSZip(this.excel.generate(jszipBuffer))
            .file('xl/sharedStrings.xml', Mustache.render(this.excel.file('xl/sharedStrings.xml').asText(), bindData))
            .generate(option);
    }

    /**
     * * _parseCommonStringWithVariable
     * * @return {Array} including common strings only having mustache-variable
     * * @private
     **/
    _parseCommonStringWithVariable(){
        let commonStringsWithVariable = [];
        _.each(this.sharedstrings,(stringObj, index)=>{
            if(_(stringObj.t).stringValue() && _(_(stringObj.t).stringValue()).hasVariable()){
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
     * * _parseDirInExcel
     * * @param {String} dir directory name in Zip file.
     * * @return {Promise|Array} array including files parsed by xml2js
     * * @private
     **/
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
     * * _buildNewSheet
     * * @param {Object} sourceSheet
     * * @param {Array} commonStringsWithVariable
     * * @return {Object}
     * * @private
     **/
    _buildNewSheet(sourceSheet, commonStringsWithVariable){
        let addedSheet = _(sourceSheet).deepCopy();
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
        addedSheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected = '0';
        return addedSheet;
    }

    /**
     * * _availableSheetid
     * * @return {String} id of next sheet
     * * @private
     **/
    _availableSheetid(){
        let maxRel = _.max(this.workbookxmlRels.Relationships.Relationship, (e)=> Number(e['$'].Id.replace('rId','')));
        let nextId = 'rId' + ('00' + (((maxRel['$'].Id.replace('rId','') >> 0))+1)).slice(-3);
        return nextId;
    }

    /**
     * * _sheetByName
     * * @param {String} sheetname target sheet name
     * * @return {Object} sheet object
     * * @private
     **/
    _sheetByName(sheetname){
        let targetSheet = _.find(this.workbookxml.workbook.sheets[0].sheet, (e)=> (e['$'].name === sheetname));
        if(!targetSheet) return null;  //invalid sheet name

        let sheetid = targetSheet['$']['r:id'];
        let targetFilePath = _.max(this.workbookxmlRels.Relationships.Relationship, (e)=>(e['$'].Id === sheetid))['$'].Target;
        let targetFileName = _.last(targetFilePath.split('/'));
        return {path: targetFilePath, value: _.find(this.sheetXmls, (e)=>(e.name === targetFileName))};
    }

    /**
     * * _sheetRelsByName
     * * @param {String} sheetname target sheet name
     * * @return {Object} sheet_rels object
     * * @private
     **/
    _sheetRelsByName(sheetname){
        let targetFilePath = this._sheetByName(sheetname).path;
        let targetName = `${_.last(targetFilePath.split('/'))}.rels`;
        return {name: targetName, value: _.find(this.sheetXmlsRels, e=>(e.name === targetName))};
    }

    /**
     * * _templateSheetRels
     * * @return {Object} sheet_rels object of template-sheet
     * * @private
     **/
    _templateSheetRels(){
        return this._sheetRelsByName(this.templateSheetName);
    }


    /**
     * * _sheetNames
     * * @return {Array} array including sheet name
     * * @private
     **/
    _sheetNames(){
        return _.map(this.sheetXmls, (e)=>e.name);
    }

    /**
     * * _cleanSharedStrings
     * * @return {Array} shared strings
     * * @private
     **/
    _cleanSharedStrings(){
        return _.map(this.sharedstrings, e => ({t:e.t, phoneticPr:e.phoneticPr}));
    }

    /**
     * * _firstSheetName
     * * @return {String} name of first-sheet of MS-Excel file
     * * @private
     **/
    _firstSheetName(){
        return this.workbookxml.workbook.sheets[0].sheet[0]['$'].name;
    }

    /**
     * * _activeSheets
     * * @return {Array} array including only active sheets.
     * * @private
     **/
    _activeSheets(){
        return _.filter(this.sheetXmls, (sheet)=>(sheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected === '1'));
    }

    /**
     * * _deactiveSheets
     * * @return {Array} array including only deactive sheets.
     * * @private
     **/
    _deactiveSheets(){
        return _.filter(this.sheetXmls, (sheet)=>(sheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected === '0'));
    }
}

//Exports
module.exports = SpreadSheet;
