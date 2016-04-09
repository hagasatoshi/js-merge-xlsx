/**
 * Excel
 * JSZip extension class
 * @author Satoshi Haga
 * @date 2016/03/27
 */

let Excel = require('jszip');
const Promise = require('bluebird');
const xml2js = require('xml2js');
const parseString = Promise.promisify(xml2js.parseString);
const _ = require('underscore');
require('./underscore_mixin');

const config = {
    FILE_SHARED_STRINGS: 'xl/sharedStrings.xml',
    FILE_WORKBOOK_RELS: 'xl/_rels/workbook.xml.rels',
    FILE_WORKBOOK: 'xl/workbook.xml',
    DIR_WORKSHEETS: 'xl/worksheets',
    DIR_WORKSHEETS_RELS: 'xl/worksheets/_rels'
};

_.extend(Excel.prototype, {

    sharedStrings: function(){
        return this.file(config.FILE_SHARED_STRINGS).asText();
    },

    parseSharedStrings: function(){
        return this.parseFile(config.FILE_SHARED_STRINGS);
    },

    hasAsSharedString: function(targetStr){
        return (this.sharedStrings().indexOf(targetStr) !== -1);
    },

    setSharedStrings: function(obj){
        if(obj){
            this.file(config.FILE_SHARED_STRINGS, _.xml(obj));
        }
        return this;
    },

    parseWorkbookRels: function(){
        return this.parseFile(config.FILE_WORKBOOK_RELS);
    },

    setWorkbookRels: function(obj){
        this.file(config.FILE_WORKBOOK_RELS, _.xml(obj));
        return this;
    },

    parseWorkbook: function(){
        return this.parseFile(config.FILE_WORKBOOK);
    },

    setWorkbook: function(obj){
        this.file(config.FILE_WORKBOOK, _.xml(obj));
        return this;
    },

    parseWorksheetsDir: function(){
        return this.parseDir(config.DIR_WORKSHEETS);
    },

    setWorksheet: function(sheetName, obj){
        this.file(`${config.DIR_WORKSHEETS}/${sheetName}`, _.xml(obj));
        return this;
    },

    setWorksheets: function(files) {
        _.each(files, ({name, data}) => {
            this.setWorksheet(name, data);
        });
        return this;
    },

    removeWorksheet: function(sheetName){
        this.remove(`${config.DIR_WORKSHEETS}/${sheetName}`);
        return this;
    },

    parseWorksheetRelsDir: function(){
        return this.parseDir(config.DIR_WORKSHEETS_RELS);
    },

    templateSheetRel: function() {
        return this.parseWorksheetRelsDir()
        .then(function(sheetXmlsRels) {
            return sheetXmlsRels ? { Relationships: sheetXmlsRels[0].Relationships } : null;
        });
    },

    setWorksheetRel: function(sheetName, obj){
        this.file(`${config.DIR_WORKSHEETS_RELS}/${sheetName}.rels`, _.xml(obj));
        return this;
    },

    setWorksheetRels: function(sheetNames, value){
        if(!value){
            return this;
        }
        let valueString = _.xml(value);
        _.each(sheetNames, (sheetName) => {
            this.file(`${config.DIR_WORKSHEETS_RELS}/${sheetName}.rels`, valueString);
        });
        return this;
    },

    removeWorksheetRel: function(sheetName){
        this.remove(`${config.DIR_WORKSHEETS_RELS}/${sheetName}.rels`);
        return this;
    },

    parseFile: function(filePath){
        return parseString(this.file(filePath).asText());
    },
    parseDir: function(dir){
        let files = this.folder(dir).file(/.xml/);
        let fileXmls = [];
        return files.reduce(
            (promise, file)=>
                promise.then((prior_file)=>
                    parseString(this.file(file.name).asText())
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
});

module.exports = Excel;