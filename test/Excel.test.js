'use strict';

var Promise = require('bluebird');
var _ = require('underscore');
var fs = Promise.promisifyAll(require('fs'));
var Excel = require('../lib/Excel');
require('../lib/underscore_mixin');
var assert = require('chai').assert;
var config = require('../lib/Config');

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
                assert.isOk(sharedStrings.includes('{{AccountName__c}}'));
                assert.isOk(sharedStrings.includes('{{AccountAddress__c}}'));
                assert.isOk(sharedStrings.includes('{{SalaryDate__c}}'));
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
                assert.isOk(sharedStrings.includes('雇用期間'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should read as encoded string', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'TemplateWithXmlEntity.xlsx').then(function (template) {
                var sharedStrings = new Excel(template).sharedStrings();
                assert.isOk(sharedStrings.includes('\&lt;\&gt;\"\\\&amp;\''));
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
                assert.isOk(sharedStrings.includes('<anyKey>anyValue</anyKey>'));
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
                assert.isOk(sharedStrings.includes('<anyKey>日本語</anyKey>'));
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
                assert.isOk(sharedStrings.includes('<anyKey>\&lt;\&gt;\"\\\&amp;\'</anyKey>'));
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
                    return e.includes('worksheets/');
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
                assert.isOk(workbookRels.includes('<anyKey>anyValue</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set Japanese value as xml string', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workbookRels = new Excel(template).setWorkbookRels({ anyKey: '日本語' }).file(config.EXCEL_FILES.FILE_WORKBOOK_RELS).asText();
                assert.isOk(workbookRels.includes('<anyKey>日本語</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set value with encoding', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workbookRels = new Excel(template).setWorkbookRels({ anyKey: '<>\"\\\&\'' }).file(config.EXCEL_FILES.FILE_WORKBOOK_RELS).asText();
                assert.isOk(workbookRels.includes('<anyKey>\&lt;\&gt;\"\\\&amp;\'</anyKey>'));
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
                assert.isOk(workbookRels.includes('<anyKey>anyValue</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set Japanese value as xml string', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workbookRels = new Excel(template).setWorkbook({ anyKey: '日本語' }).file(config.EXCEL_FILES.FILE_WORKBOOK).asText();
                assert.isOk(workbookRels.includes('<anyKey>日本語</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set value with encoding', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workbookRels = new Excel(template).setWorkbook({ anyKey: '<>\"\\\&\'' }).file(config.EXCEL_FILES.FILE_WORKBOOK).asText();
                assert.isOk(workbookRels.includes('<anyKey>\&lt;\&gt;\"\\\&amp;\'</anyKey>'));
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
                assert.isOk(workSheet.includes('<anyKey>anyValue</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set Japanese value as xml string', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workSheet = new Excel(template).setWorksheet('someSheet.xml', { anyKey: '日本語' }).file(config.EXCEL_FILES.DIR_WORKSHEETS + '/someSheet.xml').asText();
                assert.isOk(workSheet.includes('<anyKey>日本語</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set value with encoding', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workSheet = new Excel(template).setWorksheet('someSheet.xml', { anyKey: '<>\"\\\&\'' }).file(config.EXCEL_FILES.DIR_WORKSHEETS + '/someSheet.xml').asText();
                assert.isOk(workSheet.includes('<anyKey>\&lt;\&gt;\"\\\&amp;\'</anyKey>'));
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
                assert.isOk(workSheet.includes('<anyKey>anyValue</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set each value', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var excelTemplate = new Excel(template).setWorksheets([{ name: 'sheet1.xml', data: { key1: 'value1' } }, { name: 'sheet2.xml', data: { key2: 'value2' } }, { name: 'sheet3.xml', data: { key3: 'value3' } }]);
                var sheet1 = excelTemplate.file(config.EXCEL_FILES.DIR_WORKSHEETS + '/sheet1.xml').asText();
                assert.isOk(sheet1.includes('<key1>value1</key1>'));

                var sheet2 = excelTemplate.file(config.EXCEL_FILES.DIR_WORKSHEETS + '/sheet2.xml').asText();
                assert.isOk(sheet2.includes('<key2>value2</key2>'));

                var sheet3 = excelTemplate.file(config.EXCEL_FILES.DIR_WORKSHEETS + '/sheet3.xml').asText();
                assert.isOk(sheet3.includes('<key3>value3</key3>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set Japanese value as xml string', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workSheet = new Excel(template).setWorksheets([{ name: 'sheet1.xml', data: { anyKey: '日本語' } }]).file(config.EXCEL_FILES.DIR_WORKSHEETS + '/sheet1.xml').asText();
                assert.isOk(workSheet.includes('<anyKey>日本語</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should set value with encoding', function () {
            return fs.readFileAsync(config.TEST_DIRS.TEMPLATE + 'Template.xlsx').then(function (template) {
                var workSheet = new Excel(template).setWorksheets([{ name: 'sheet1.xml', data: { anyKey: '<>\"\\\&\'' } }]).file(config.EXCEL_FILES.DIR_WORKSHEETS + '/sheet1.xml').asText();
                assert.isOk(workSheet.includes('<anyKey>\&lt;\&gt;\"\\\&amp;\'</anyKey>'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });
    });
});