/**
 * Excel
 * @author Satoshi Haga
 * @date 2016/03/27
 */

const Mustache = require('mustache');
const Promise = require('bluebird');
const parseString = Promise.promisify(require('xml2js').parseString);
const _ = require('underscore');
require('./underscore_mixin');
const config = require('./Config');
let Excel = require('jszip');

_.extend(Excel.prototype, {

    //read as encoded strings
    sharedStrings: function() {
        return this.file(config.EXCEL_FILES.FILE_SHARED_STRINGS).asText();
    },

    variables: function() {
        return _.variables(this.sharedStrings());
    },

    parseSharedStrings: function() {
        return this.parseFile(config.EXCEL_FILES.FILE_SHARED_STRINGS);
    },

    //match as encoded strings
    hasAsSharedString: function(targetStr) {
        return (this.sharedStrings().indexOf(targetStr) !== -1);
    },

    //save with xml-encoding
    setSharedStrings: function(obj) {
        if(obj) {
            this.file(config.EXCEL_FILES.FILE_SHARED_STRINGS, _.xml(obj));
        }
        return this;
    },

    parseWorkbookRels: function() {
        return this.parseFile(config.EXCEL_FILES.FILE_WORKBOOK_RELS);
    },

    setWorkbookRels: function(obj) {
        this.file(config.EXCEL_FILES.FILE_WORKBOOK_RELS, _.xml(obj));
        return this;
    },

    parseWorkbook: function() {
        return this.parseFile(config.EXCEL_FILES.FILE_WORKBOOK);
    },

    setWorkbook: function(obj) {
        this.file(config.EXCEL_FILES.FILE_WORKBOOK, _.xml(obj));
        return this;
    },

    parseWorksheetsDir: function() {
        return this.parseDir(config.EXCEL_FILES.DIR_WORKSHEETS);
    },

    setWorksheet: function(sheetName, obj) {
        this.file(`${config.EXCEL_FILES.DIR_WORKSHEETS}/${sheetName}`, _.xml(obj));
        return this;
    },

    setWorksheets: function(files) {
        _.each(files, ({name, data}) => {
            this.setWorksheet(name, data);
        });
        return this;
    },

    removeWorksheet: function(sheetName) {
        this.remove(`${config.EXCEL_FILES.DIR_WORKSHEETS}/${sheetName}`);
        this.remove(`${config.EXCEL_FILES.DIR_WORKSHEETS_RELS}/${sheetName}.rels`);
        return this;
    },

    parseWorksheetRelsDir: function() {
        return this.parseDir(config.EXCEL_FILES.DIR_WORKSHEETS_RELS);
    },

    setTemplateSheetRel: function() {
        return this.parseWorksheetRelsDir()
        .then((sheetXmlsRels) => {
            this.templateSheetRel = sheetXmlsRels ?
                {Relationships: sheetXmlsRels[0].Relationships} : null;
            return this;
        });
    },

    setWorksheetRel: function(sheetName, obj) {
        this.file(`${config.EXCEL_FILES.DIR_WORKSHEETS_RELS}/${sheetName}.rels`, _.xml(obj));
        return this;
    },

    setWorksheetRels: function(sheetNames) {
        if(!this.templateSheetRel) {
            return this;
        }
        let valueString = _.xml(this.templateSheetRel);
        _.each(sheetNames, (sheetName) => {
            this.file(`${config.EXCEL_FILES.DIR_WORKSHEETS_RELS}/${sheetName}.rels`, valueString);
        });
        return this;
    },

    parseFile: function(filePath) {
        return parseString(this.file(filePath).asText());
    },

    parseDir: function(dir) {
        let files = this.folder(dir).file(/.xml/);
        let fileXmls = [];
        return files.reduce(
            (promise, file) =>
                promise.then((prior_file) =>
                    parseString(this.file(file.name).asText())
                        .then((file_xml) => {
                            file_xml.name = _.last(file.name.split('/'));
                            fileXmls.push(file_xml);
                            return fileXmls;
                        })
                )
            ,
            Promise.resolve()
        );
    },

    merge: function(mergedData) {
        return this.file('xl/sharedStrings.xml', Mustache.render(this.sharedStrings(), mergedData))
    },

    generateWithData(excelObj) {
        return this.setTemplateSheetRel()
            .then(() => {
                return this.setSharedStrings(excelObj.sharedstrings.value())
                    .setWorkbookRels(excelObj.relationship.value())
                    .setWorkbook(excelObj.workbookxml.value())
                    .setWorksheets(excelObj.sheetXmls.value())
                    .setWorksheetRels(excelObj.sheetXmls.names())
                    .generate({
                        type:        config.JSZIP_OPTION.BUFFER_TYPE_OUTPUT,
                        compression: config.JSZIP_OPTION.COMPLESSION}
                    );
            })
    }
});

module.exports = Excel;