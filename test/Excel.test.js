'use strict';

var Promise = require('bluebird');
var _ = require('underscore');
var fs = Promise.promisifyAll(require('fs'));
var Excel = require('../lib/Excel');
require('../lib/underscore');
var assert = require('chai').assert;
var config = require('../lib/Config');
var xml2js = require('xml2js');
var builder = new xml2js.Builder();

var readFiles = function readFiles(template) {
    return Promise.props({
        template: fs.readFileAsync('' + config.TEST_DIRS.TEMPLATE + template)
    });
};

describe('Excel.js', function () {
    describe('sharedStrings()', function () {

        it('should read each strings on template', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var sharedStrings = new Excel(template).sharedStrings();
                assert.isOk(typeof sharedStrings === 'string');
                assert.isOk(_.includeString(sharedStrings, '{{AccountName__c}}'));
                assert.isOk(_.includeString(sharedStrings, '{{AccountAddress__c}}'));
                assert.isOk(_.includeString(sharedStrings, '{{SalaryDate__c}}'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should read with no error in case no strings defined', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'TemplateNoStrings.xlsx').then(function (template) {
                var sharedStrings = new Excel(template).sharedStrings();
                assert.isOk(typeof sharedStrings === 'string');
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should read Japanese strings on template', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var sharedStrings = new Excel(template).sharedStrings();
                assert.isOk(_.includeString(sharedStrings, '雇用期間'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should read as encoded string', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'TemplateWithXmlEntity.xlsx').then(function (template) {
                var sharedStrings = new Excel(template).sharedStrings();
                assert.isOk(_.includeString(sharedStrings, '\&lt;\&gt;\"\\\&amp;\''));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });
    });
    describe('parseSharedStrings()', function () {

        it('should parse each strings on template', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                return new Excel(template).parseSharedStrings();
            }).then(function (stringModels) {
                var si = stringModels.sst.si;
                assert.notStrictEqual(si, undefined);
                assert.isOk(si instanceof Array);

                si = _.map(si, function (e) {
                    return _.stringValue(e.t);
                });
                assert.isOk(_.containsAsPartialString(si, '{{AccountName__c}}'));
                assert.isOk(_.containsAsPartialString(si, '{{AccountAddress__c}}'));
                assert.isOk(_.containsAsPartialString(si, '{{SalaryDate__c}}'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should parse with no error in case no strings defined', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'TemplateNoStrings.xlsx').then(function (template) {
                return new Excel(template).parseSharedStrings();
            }).then(function (stringModels) {
                assert.isOk(!stringModels.sst || !stringModels.sst.si);
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should parse Japanese with no error', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                return new Excel(template).parseSharedStrings();
            }).then(function (stringModels) {
                var si = stringModels.sst.si;
                assert.notStrictEqual(si, undefined);
                assert.isOk(si instanceof Array);

                si = _.map(si, function (e) {
                    return _.stringValue(e.t);
                });
                assert.isOk(_.containsAsPartialString(si, '雇用期間'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should parse as decoded string', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'TemplateWithXmlEntity.xlsx').then(function (template) {
                return new Excel(template).parseSharedStrings();
            }).then(function (stringModels) {
                var si = stringModels.sst.si;
                assert.notStrictEqual(si, undefined);
                assert.isOk(si instanceof Array);

                si = _.map(si, function (e) {
                    return _.stringValue(e.t);
                });
                assert.isOk(_.containsAsPartialString(si, '<>\"\\\&\''));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });
    });

    describe('setSharedStrings()', function () {

        it('should return this instance', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var test = new Excel(template).setSharedStrings();
                assert.notStrictEqual(test, undefined);
                assert.notStrictEqual(test, null);
                assert.isOk(test instanceof Excel);
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set value as xml string', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var sharedStrings = new Excel(template).setSharedStrings({
                    anyKey: 'anyValue'
                }).sharedStrings();
                assert.isOk(_.includeString(sharedStrings, '<anyKey>anyValue</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set Japanese value as xml string', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var sharedStrings = new Excel(template).setSharedStrings({
                    anyKey: '日本語'
                }).sharedStrings();
                assert.isOk(_.includeString(sharedStrings, '<anyKey>日本語</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set value with encoding', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var sharedStrings = new Excel(template).setSharedStrings({
                    anyKey: '<>\"\\\&\''
                }).sharedStrings();
                assert.isOk(_.includeString(sharedStrings, '<anyKey>\&lt;\&gt;\"\\\&amp;\'</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });
    });

    describe('parseWorkbookRels()', function () {

        it('should parse relation files, styles/sharedStrings/worksheets/theme', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                return new Excel(template).parseWorkbookRels();
            }).then(function (workbookRels) {
                var relationships = workbookRels.Relationships.Relationship;
                relationships = _.map(relationships, function (e) {
                    return e['$'].Target;
                });
                assert.isOk(_.containsAsPartialString(relationships, 'styles.xml'));
                assert.isOk(_.containsAsPartialString(relationships, 'sharedStrings.xml'));
                assert.isOk(_.containsAsPartialString(relationships, 'worksheets/'));
                assert.isOk(_.containsAsPartialString(relationships, 'theme/'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should parse each relation', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template3Sheet.xlsx').then(function (template) {
                return new Excel(template).parseWorkbookRels();
            }).then(function (workbookRels) {
                var sheetCount = _.chain(workbookRels.Relationships.Relationship).map(function (e) {
                    return e['$'].Target;
                }).filter(function (e) {
                    return _.includeString(e, 'worksheets/');
                }).value().length;
                assert.strictEqual(sheetCount, 3);
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });
    });

    describe('setWorkbookRels()', function () {

        it('should return this instance', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var test = new Excel(template).setWorkbookRels({});
                assert.notStrictEqual(test, undefined);
                assert.notStrictEqual(test, null);
                assert.isOk(test instanceof Excel);
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set value as xml string', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workbookRels = new Excel(template).setWorkbookRels({ anyKey: 'anyValue' }).file(config.EXCEL_FILES.FILE_WORKBOOK_RELS).asText();
                assert.isOk(_.includeString(workbookRels, '<anyKey>anyValue</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set Japanese value as xml string', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workbookRels = new Excel(template).setWorkbookRels({ anyKey: '日本語' }).file(config.EXCEL_FILES.FILE_WORKBOOK_RELS).asText();
                assert.isOk(_.includeString(workbookRels, '<anyKey>日本語</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set value with encoding', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workbookRels = new Excel(template).setWorkbookRels({ anyKey: '<>\"\\\&\'' }).file(config.EXCEL_FILES.FILE_WORKBOOK_RELS).asText();
                assert.isOk(_.includeString(workbookRels, '<anyKey>\&lt;\&gt;\"\\\&amp;\'</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });
    });

    describe('parseWorkbook()', function () {

        it('should parse information of sheet', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                return new Excel(template).parseWorkbook();
            }).then(function (workbook) {
                var sheets = workbook.workbook.sheets[0].sheet;
                assert.notStrictEqual(sheets, undefined);
                assert.notStrictEqual(sheets, null);
                assert.strictEqual(sheets.length, 1);
                assert.strictEqual(sheets[0]['$'].name, 'Sheet1');
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should parse each sheet', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template3Sheet.xlsx').then(function (template) {
                return new Excel(template).parseWorkbook();
            }).then(function (workbook) {
                var sheets = workbook.workbook.sheets[0].sheet;
                assert.notStrictEqual(sheets, undefined);
                assert.notStrictEqual(sheets, null);
                assert.strictEqual(sheets.length, 3);
                assert.strictEqual(sheets[0]['$'].name, 'Sheet1');
                assert.strictEqual(sheets[1]['$'].name, 'Sheet2');
                assert.strictEqual(sheets[2]['$'].name, 'Sheet3');
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });
    });

    describe('setWorkbook()', function () {

        it('should return this instance', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var test = new Excel(template).setWorkbook({});
                assert.notStrictEqual(test, undefined);
                assert.notStrictEqual(test, null);
                assert.isOk(test instanceof Excel);
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set value as xml string', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workbookRels = new Excel(template).setWorkbook({ anyKey: 'anyValue' }).file(config.EXCEL_FILES.FILE_WORKBOOK).asText();
                assert.isOk(_.includeString(workbookRels, '<anyKey>anyValue</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set Japanese value as xml string', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workbookRels = new Excel(template).setWorkbook({ anyKey: '日本語' }).file(config.EXCEL_FILES.FILE_WORKBOOK).asText();
                assert.isOk(_.includeString(workbookRels, '<anyKey>日本語</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set value with encoding', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workbookRels = new Excel(template).setWorkbook({ anyKey: '<>\"\\\&\'' }).file(config.EXCEL_FILES.FILE_WORKBOOK).asText();
                assert.isOk(_.includeString(workbookRels, '<anyKey>\&lt;\&gt;\"\\\&amp;\'</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });
    });

    describe('parseWorksheetsDir()', function () {

        it('should parse relation and contents', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                return new Excel(template).parseWorksheetsDir();
            }).then(function (worksheets) {
                var files = _.filter(worksheets, function (e) {
                    return !!e.worksheet;
                });
                assert.strictEqual(files.length, 1);

                var relations = _.filter(worksheets, function (e) {
                    return !!e.Relationships;
                });
                assert.strictEqual(relations.length, 1);

                assert.strictEqual(files[0].name + '.rels', relations[0].name);
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should parse each relation and contents', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template3Sheet.xlsx').then(function (template) {
                return new Excel(template).parseWorksheetsDir();
            }).then(function (worksheets) {
                var files = _.filter(worksheets, function (e) {
                    return !!e.worksheet;
                });
                assert.strictEqual(files.length, 3);

                var relations = _.filter(worksheets, function (e) {
                    return !!e.Relationships;
                });
                assert.strictEqual(relations.length, 3);

                var fileNameInRelations = _.map(relations, function (e) {
                    return e.name;
                });
                assert.isOk(_.contains(fileNameInRelations, files[0].name + '.rels'));
                assert.isOk(_.contains(fileNameInRelations, files[1].name + '.rels'));
                assert.isOk(_.contains(fileNameInRelations, files[2].name + '.rels'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });
    });

    describe('setWorksheet()', function () {

        it('should return this instance', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var test = new Excel(template).setWorksheet('someSheet.xml', {});
                assert.notStrictEqual(test, undefined);
                assert.notStrictEqual(test, null);
                assert.isOk(test instanceof Excel);
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set value as xml string', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workSheet = new Excel(template).setWorksheet('someSheet.xml', { anyKey: 'anyValue' }).file(config.EXCEL_FILES.DIR_WORKSHEETS + '/someSheet.xml').asText();
                assert.isOk(_.includeString(workSheet, '<anyKey>anyValue</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set Japanese value as xml string', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workSheet = new Excel(template).setWorksheet('someSheet.xml', { anyKey: '日本語' }).file(config.EXCEL_FILES.DIR_WORKSHEETS + '/someSheet.xml').asText();
                assert.isOk(_.includeString(workSheet, '<anyKey>日本語</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set value with encoding', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workSheet = new Excel(template).setWorksheet('someSheet.xml', { anyKey: '<>\"\\\&\'' }).file(config.EXCEL_FILES.DIR_WORKSHEETS + '/someSheet.xml').asText();
                assert.isOk(_.includeString(workSheet, '<anyKey>\&lt;\&gt;\"\\\&amp;\'</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });
    });

    describe('setWorksheets()', function () {

        it('should return this instance', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var test = new Excel(template).setWorksheets([]);
                assert.notStrictEqual(test, undefined);
                assert.notStrictEqual(test, null);
                assert.isOk(test instanceof Excel);
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set value as xml string', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workSheet = new Excel(template).setWorksheets([{ name: 'sheet1.xml', data: { anyKey: 'anyValue' } }]).file(config.EXCEL_FILES.DIR_WORKSHEETS + '/sheet1.xml').asText();
                assert.isOk(_.includeString(workSheet, '<anyKey>anyValue</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set each value', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var excelTemplate = new Excel(template).setWorksheets([{ name: 'sheet1.xml', data: { key1: 'value1' } }, { name: 'sheet2.xml', data: { key2: 'value2' } }, { name: 'sheet3.xml', data: { key3: 'value3' } }]);
                var sheet1 = excelTemplate.file(config.EXCEL_FILES.DIR_WORKSHEETS + '/sheet1.xml').asText();
                assert.isOk(_.includeString(sheet1, '<key1>value1</key1>'));

                var sheet2 = excelTemplate.file(config.EXCEL_FILES.DIR_WORKSHEETS + '/sheet2.xml').asText();
                assert.isOk(_.includeString(sheet2, '<key2>value2</key2>'));

                var sheet3 = excelTemplate.file(config.EXCEL_FILES.DIR_WORKSHEETS + '/sheet3.xml').asText();
                assert.isOk(_.includeString(sheet3, '<key3>value3</key3>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set Japanese value as xml string', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workSheet = new Excel(template).setWorksheets([{ name: 'sheet1.xml', data: { anyKey: '日本語' } }]).file(config.EXCEL_FILES.DIR_WORKSHEETS + '/sheet1.xml').asText();
                assert.isOk(_.includeString(workSheet, '<anyKey>日本語</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set value with encoding', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workSheet = new Excel(template).setWorksheets([{ name: 'sheet1.xml', data: { anyKey: '<>\"\\\&\'' } }]).file(config.EXCEL_FILES.DIR_WORKSHEETS + '/sheet1.xml').asText();
                assert.isOk(_.includeString(workSheet, '<anyKey>\&lt;\&gt;\"\\\&amp;\'</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });
    });

    describe('removeWorksheet()', function () {

        it('should return this instance', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var templateObj = new Excel(template).setWorksheets([{ name: 'sheet1.xml', data: { anyKey: 'anyValue' } }]);
                var test = templateObj.removeWorksheet('sheet1.xml');
                assert.notStrictEqual(test, undefined);
                assert.notStrictEqual(test, null);
                assert.isOk(test instanceof Excel);
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should remove the file', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var templateObj = new Excel(template).setWorksheets([{ name: 'sheet1.xml', data: { anyKey: 'anyValue' } }]);
                templateObj.removeWorksheet('sheet1.xml');
                assert.strictEqual(templateObj.file(config.EXCEL_FILES.DIR_WORKSHEETS + '/sheet1.xml'), null);
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should not remove other files', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var templateObj = new Excel(template).setWorksheets([{ name: 'sheet1.xml', data: { anyKey: 'anyValue' } }, { name: 'sheet2.xml', data: { anyKey: 'anyValue' } }, { name: 'sheet3.xml', data: { anyKey: 'anyValue' } }]);
                templateObj.removeWorksheet('sheet1.xml');
                assert.notStrictEqual(templateObj.file(config.EXCEL_FILES.DIR_WORKSHEETS + '/sheet2.xml'), null);
                assert.notStrictEqual(templateObj.file(config.EXCEL_FILES.DIR_WORKSHEETS + '/sheet3.xml'), null);
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should return with no error even if not existing', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var templateObj = new Excel(template);
                var test = templateObj.removeWorksheet('invalid sheet name');
                assert.isOk(test instanceof Excel);
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });
    });

    describe('parseWorksheetRelsDir()', function () {

        it('should parse relation file', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                return new Excel(template).parseWorksheetRelsDir();
            }).then(function (worksheetRels) {
                assert.notStrictEqual(worksheetRels, undefined);
                assert.notStrictEqual(worksheetRels, null);
                assert.isOk(_.isArray(worksheetRels));
                assert.strictEqual(1, worksheetRels.length);
                assert.isOk(_.consistOf(worksheetRels, ['Relationships', 'name']));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should parse each relation file', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template3Sheet.xlsx').then(function (template) {
                return new Excel(template).parseWorksheetRelsDir();
            }).then(function (worksheetRels) {
                assert.notStrictEqual(worksheetRels, undefined);
                assert.notStrictEqual(worksheetRels, null);
                assert.isOk(_.isArray(worksheetRels));
                assert.strictEqual(3, worksheetRels.length);
                assert.isOk(_.consistOf(worksheetRels, ['Relationships', 'name']));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });
    });

    describe('setTemplateSheetRel()', function () {

        it('should return this instance', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                return new Excel(template).setTemplateSheetRel();
            }).then(function (templateObj) {
                assert.notStrictEqual(templateObj, undefined);
                assert.notStrictEqual(templateObj, null);
                assert.isOk(templateObj instanceof Excel);
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set relation file as template sheet', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                return new Excel(template).setTemplateSheetRel();
            }).then(function (templateObj) {
                assert.notStrictEqual(templateObj.templateSheetRel, undefined);
                assert.notStrictEqual(templateObj.templateSheetRel.Relationships, undefined);
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });
    });

    describe('setWorksheetRel()', function () {

        it('should return this instance', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var test = new Excel(template).setWorksheetRel('someSheet.xml', {});
                assert.notStrictEqual(test, undefined);
                assert.notStrictEqual(test, null);
                assert.isOk(test instanceof Excel);
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set value as xml string', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workSheetRels = new Excel(template).setWorksheetRel('someSheet.xml', { anyKey: 'anyValue' }).file(config.EXCEL_FILES.DIR_WORKSHEETS_RELS + '/someSheet.xml.rels').asText();
                assert.isOk(_.includeString(workSheetRels, '<anyKey>anyValue</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set Japanese value as xml string', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workSheetRels = new Excel(template).setWorksheetRel('someSheet.xml', { anyKey: '日本語' }).file(config.EXCEL_FILES.DIR_WORKSHEETS_RELS + '/someSheet.xml.rels').asText();
                assert.isOk(_.includeString(workSheetRels, '<anyKey>日本語</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set value with encoding', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workSheetRels = new Excel(template).setWorksheetRel('someSheet.xml', { anyKey: '<>\"\\\&\'' }).file(config.EXCEL_FILES.DIR_WORKSHEETS_RELS + '/someSheet.xml.rels').asText();
                assert.isOk(_.includeString(workSheetRels, '<anyKey>\&lt;\&gt;\"\\\&amp;\'</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });
    });

    describe('setWorksheetRels()', function () {

        it('should return this instance if template is not set', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var test = new Excel(template).setWorksheetRels(['someSheet']);
                assert.notStrictEqual(test, undefined);
                assert.notStrictEqual(test, null);
                assert.isOk(test instanceof Excel);
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('relation file should be the same with template', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                return new Excel(template).setTemplateSheetRel();
            }).then(function (templateObj) {
                var test = templateObj.setWorksheetRels(['someSheet']);
                var workSheetRels = templateObj.setWorksheetRels(['someSheet']).file(config.EXCEL_FILES.DIR_WORKSHEETS_RELS + '/someSheet.rels').asText();
                var templateString = builder.buildObject(templateObj.templateSheetRel);
                assert.strictEqual(workSheetRels, templateString);
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('all relation files should be the same with template', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                return new Excel(template).setTemplateSheetRel();
            }).then(function (templateObj) {
                templateObj = templateObj.setWorksheetRels(['someSheet1', 'someSheet2', 'someSheet3']);
                var templateString = builder.buildObject(templateObj.templateSheetRel);
                _.each(['someSheet1', 'someSheet2', 'someSheet3'], function (sheetName) {
                    var sheetStr = templateObj.file(config.EXCEL_FILES.DIR_WORKSHEETS_RELS + '/' + sheetName + '.rels').asText();
                    assert.strictEqual(sheetStr, templateString);
                });
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });
    });

    describe('parseFile()', function () {

        it('should parse async by returning Promise', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var promise = new Excel(template).parseFile(config.EXCEL_FILES.FILE_SHARED_STRINGS);
                assert.notStrictEqual(promise, undefined);
                assert.notStrictEqual(promise, null);
                assert.isOk(promise instanceof Promise);
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should parse from xml string', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                return new Excel(template).parseFile(config.EXCEL_FILES.FILE_SHARED_STRINGS);
            }).then(function (stringModel) {
                assert.isNotOk(stringModel instanceof String);
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });
    });

    describe('parseDir()', function () {

        it('should parse async by returning Promise', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var promise = new Excel(template).parseDir(config.EXCEL_FILES.DIR_WORKSHEETS);
                assert.notStrictEqual(promise, undefined);
                assert.notStrictEqual(promise, null);
                assert.isOk(promise instanceof Promise);
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should parse each file in the directory', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                return new Excel(template).parseDir(config.EXCEL_FILES.DIR_WORKSHEETS);
            }).then(function (fileModels) {
                assert.isOk(_.consistOf(fileModels, ['name']));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });
    });
});