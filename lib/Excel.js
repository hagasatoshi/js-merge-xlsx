/**
 * Excel
 * @author Satoshi Haga
 * @date 2016/03/27
 */

'use strict';

var Promise = require('bluebird');
var xml2js = require('xml2js');
var parseString = Promise.promisify(xml2js.parseString);
var builder = new xml2js.Builder();
var _ = require('underscore');
require('./underscore');
var config = require('./Config');
var Excel = require('jszip');

_.extend(Excel.prototype, {

    //read as encoded strings
    sharedStrings: function sharedStrings() {
        return this.file(config.EXCEL_FILES.FILE_SHARED_STRINGS).asText();
    },

    parseSharedStrings: function parseSharedStrings() {
        return this.parseFile(config.EXCEL_FILES.FILE_SHARED_STRINGS);
    },

    //save with xml-encoding
    setSharedStrings: function setSharedStrings(obj) {
        if (obj) {
            this.file(config.EXCEL_FILES.FILE_SHARED_STRINGS, builder.buildObject(obj));
        }
        return this;
    },

    parseWorkbookRels: function parseWorkbookRels() {
        return this.parseFile(config.EXCEL_FILES.FILE_WORKBOOK_RELS);
    },

    setWorkbookRels: function setWorkbookRels(obj) {
        this.file(config.EXCEL_FILES.FILE_WORKBOOK_RELS, builder.buildObject(obj));
        return this;
    },

    parseWorkbook: function parseWorkbook() {
        return this.parseFile(config.EXCEL_FILES.FILE_WORKBOOK);
    },

    setWorkbook: function setWorkbook(obj) {
        this.file(config.EXCEL_FILES.FILE_WORKBOOK, builder.buildObject(obj));
        return this;
    },

    parseWorksheetsDir: function parseWorksheetsDir() {
        return this.parseDir(config.EXCEL_FILES.DIR_WORKSHEETS);
    },

    setWorksheet: function setWorksheet(sheetName, obj) {
        this.file(config.EXCEL_FILES.DIR_WORKSHEETS + '/' + sheetName, builder.buildObject(obj));
        return this;
    },

    setWorksheets: function setWorksheets(files) {
        var _this = this;

        _.each(files, function (_ref) {
            var name = _ref.name;
            var data = _ref.data;

            _this.setWorksheet(name, data);
        });
        return this;
    },

    removeWorksheet: function removeWorksheet(sheetName) {
        var filePath = config.EXCEL_FILES.DIR_WORKSHEETS + '/' + sheetName;
        var relsFilePath = filePath + '.rels';
        if (!this.file(filePath)) {
            return this;
        }
        this.remove(filePath);
        this.remove(relsFilePath);
        return this;
    },

    parseWorksheetRelsDir: function parseWorksheetRelsDir() {
        return this.parseDir(config.EXCEL_FILES.DIR_WORKSHEETS_RELS);
    },

    setTemplateSheetRel: function setTemplateSheetRel() {
        var _this2 = this;

        return this.parseWorksheetRelsDir().then(function (sheetXmlsRels) {
            _this2.templateSheetRel = sheetXmlsRels ? { Relationships: sheetXmlsRels[0].Relationships } : null;
            return _this2;
        });
    },

    setWorksheetRel: function setWorksheetRel(sheetName, obj) {
        this.file(config.EXCEL_FILES.DIR_WORKSHEETS_RELS + '/' + sheetName + '.rels', builder.buildObject(obj));
        return this;
    },

    setWorksheetRels: function setWorksheetRels(sheetNames) {
        var _this3 = this;

        if (!this.templateSheetRel) {
            return this;
        }
        var valueString = builder.buildObject(this.templateSheetRel);
        _.each(sheetNames, function (sheetName) {
            _this3.file(config.EXCEL_FILES.DIR_WORKSHEETS_RELS + '/' + sheetName + '.rels', valueString);
        });
        return this;
    },

    parseFile: function parseFile(filePath) {
        return parseString(this.file(filePath).asText());
    },

    parseDir: function parseDir(dir) {
        var _this4 = this;

        var files = this.folder(dir).file(/.xml/);
        var fileXmls = [];
        return files.reduce(function (promise, file) {
            return promise.then(function (prior_file) {
                return parseString(_this4.file(file.name).asText()).then(function (file_xml) {
                    file_xml.name = _.last(file.name.split('/'));
                    fileXmls.push(file_xml);
                    return fileXmls;
                });
            });
        }, Promise.resolve());
    },

    generateWithData: function generateWithData(excelObj) {
        var _this5 = this;

        return this.setTemplateSheetRel().then(function () {
            return _this5.setSharedStrings(excelObj.sharedstrings.value()).setWorkbookRels(excelObj.relationship.value()).setWorkbook(excelObj.workbookxml.value()).setWorksheets(excelObj.sheetXmls.value()).setWorksheetRels(excelObj.sheetXmls.names()).generate({
                type: config.JSZIP_OPTION.BUFFER_TYPE_OUTPUT,
                compression: config.JSZIP_OPTION.COMPLESSION });
        });
    }
});

module.exports = Excel;