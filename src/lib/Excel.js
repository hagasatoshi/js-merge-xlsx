/**
 * Excel
 * @author Satoshi Haga
 * @date 2016/03/27
 */

const Mustache = require('mustache');
const Promise = require('bluebird');
const xml2js = require('xml2js');
const parseString = Promise.promisify(xml2js.parseString);
const builder = new xml2js.Builder();
const _ = require('underscore');
require('./underscore_mixin');
const config = require('./Config');
let Excel = require('jszip');

_.extend(Excel.prototype, {

    //read as encoded strings
    sharedStrings: function() {
        return this.file(config.EXCEL_FILES.FILE_SHARED_STRINGS).asText();
    },

    parseSharedStrings: function() {
        return this.parseFile(config.EXCEL_FILES.FILE_SHARED_STRINGS);
    },

    //save with xml-encoding
    setSharedStrings: function(obj) {
        if(obj) {
            this.file(config.EXCEL_FILES.FILE_SHARED_STRINGS, builder.buildObject(obj));
        }
        return this;
    },

    parseWorkbookRels: function() {
        return this.parseFile(config.EXCEL_FILES.FILE_WORKBOOK_RELS);
    },

    setWorkbookRels: function(obj) {
        this.file(config.EXCEL_FILES.FILE_WORKBOOK_RELS, builder.buildObject(obj));
        return this;
    },

    parseWorkbook: function() {
        return this.parseFile(config.EXCEL_FILES.FILE_WORKBOOK);
    },

    setWorkbook: function(obj) {
        this.file(config.EXCEL_FILES.FILE_WORKBOOK, builder.buildObject(obj));
        return this;
    },

    parseWorksheetsDir: function() {
        return this.parseDir(config.EXCEL_FILES.DIR_WORKSHEETS);
    },

    setWorksheet: function(sheetName, obj) {
        this.file(`${config.EXCEL_FILES.DIR_WORKSHEETS}/${sheetName}`, builder.buildObject(obj));
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
        this.file(
            `${config.EXCEL_FILES.DIR_WORKSHEETS_RELS}/${sheetName}.rels`,
            builder.buildObject(obj)
        );
        return this;
    },

    setWorksheetRels: function(sheetNames) {
        if(!this.templateSheetRel) {
            return this;
        }
        let valueString = builder.buildObject(this.templateSheetRel);
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