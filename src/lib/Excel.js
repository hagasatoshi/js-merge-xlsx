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
        this.file(config.FILE_SHARED_STRINGS, _.xml(obj));
    },

    parseWorkbookRels: function(){
        return this.parseFile(config.FILE_WORKBOOK_RELS);
    },

    setWorkbookRels: function(obj){
        this.file(config.FILE_WORKBOOK_RELS, _.xml(obj));
    },

    parseWorkbook: function(){
        return this.parseFile(config.FILE_WORKBOOK);
    },

    setWorkbook: function(obj){
        this.file(config.FILE_WORKBOOK, _.xml(obj));
    },

    parseWorksheetsDir: function(){
        return this.parseDir(config.DIR_WORKSHEETS);
    },

    setWorksheet: function(sheetName, obj){
        this.file(`${config.DIR_WORKSHEETS}/${sheetName}`, _.xml(obj));
    },

    removeWorksheet: function(sheetName){
        this.remove(`${config.DIR_WORKSHEETS}/${sheetName}`);
    },

    parseWorksheetRelsDir: function(){
        return this.parseDir(config.DIR_WORKSHEETS_RELS);
    },

    setWorksheetRel: function(sheetName, obj){
        this.file(`${config.DIR_WORKSHEETS_RELS}/${sheetName}.rels`, _.xml(obj));
    },

    removeWorksheetRel: function(sheetName){
        this.remove(`${config.DIR_WORKSHEETS_RELS}/${sheetName}.rels`);
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