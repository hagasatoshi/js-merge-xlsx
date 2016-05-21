const Promise = require('bluebird');
const _ = require('underscore');
const fs = Promise.promisifyAll(require('fs'));
const Excel = require('../lib/Excel');
require('../lib/underscore_mixin');
const assert = require('chai').assert;
const config = require('../lib/Config');

const readFiles = (template) => {
    return Promise.props({
        template: fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}${template}`)
    });
};

describe('Excel.js', () => {
    describe('sharedStrings()', () => {

        it('should read each strings on template', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let sharedStrings = new Excel(template).sharedStrings();
                    assert.isOk(typeof sharedStrings === 'string');
                    assert.isOk(sharedStrings.includes('{{AccountName__c}}'));
                    assert.isOk(sharedStrings.includes('{{AccountAddress__c}}'));
                    assert.isOk(sharedStrings.includes('{{SalaryDate__c}}'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should read with no error in case no strings defined', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}TemplateNoStrings.xlsx`)
                .then((template) => {
                    let sharedStrings = new Excel(template).sharedStrings();
                    assert.isOk(typeof sharedStrings === 'string');
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should read Japanese strings on template', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let sharedStrings = new Excel(template).sharedStrings();
                    assert.isOk(sharedStrings.includes('雇用期間'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should read as encoded string', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}TemplateWithXmlEntity.xlsx`)
                .then((template) => {
                    let sharedStrings = new Excel(template).sharedStrings();
                    assert.isOk(sharedStrings.includes('\&lt;\&gt;\"\\\&amp;\''));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

    });
    describe('parseSharedStrings()', () => {

        it('should parse each strings on template', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    return new Excel(template).parseSharedStrings();
                }).then((stringModels) => {
                    let si = stringModels.sst.si;
                    assert.notStrictEqual(si, undefined);
                    assert.isOk(si instanceof Array);

                    si = _.map(si, (e) => _.stringValue(e.t));
                    assert.isOk(_.containsAsPartialString(si, '{{AccountName__c}}'));
                    assert.isOk(_.containsAsPartialString(si, '{{AccountAddress__c}}'));
                    assert.isOk(_.containsAsPartialString(si, '{{SalaryDate__c}}'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should parse with no error in case no strings defined', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}TemplateNoStrings.xlsx`)
                .then((template) => {
                    return new Excel(template).parseSharedStrings();
                }).then((stringModels) => {
                    assert.isOk(!stringModels.sst || !stringModels.sst.si);
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should parse Japanese with no error', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    return new Excel(template).parseSharedStrings();
                }).then((stringModels) => {
                    let si = stringModels.sst.si;
                    assert.notStrictEqual(si, undefined);
                    assert.isOk(si instanceof Array);

                    si = _.map(si, (e) => _.stringValue(e.t));
                    assert.isOk(_.containsAsPartialString(si, '雇用期間'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should parse as decoded string', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}TemplateWithXmlEntity.xlsx`)
                .then((template) => {
                    return new Excel(template).parseSharedStrings();
                }).then((stringModels) => {
                    let si = stringModels.sst.si;
                    assert.notStrictEqual(si, undefined);
                    assert.isOk(si instanceof Array);

                    si = _.map(si, (e) => _.stringValue(e.t));
                    assert.isOk(_.containsAsPartialString(si, '<>\"\\\&\''));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

    });

    describe('setSharedStrings()', () => {

        it('should return this instance', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let test = new Excel(template).setSharedStrings();
                    assert.notStrictEqual(test, undefined);
                    assert.notStrictEqual(test, null);
                    assert.isOk(test instanceof Excel);
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should set value as xml string', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let sharedStrings = new Excel(template).setSharedStrings({
                        anyKey: 'anyValue'
                    }).sharedStrings();
                    assert.isOk(sharedStrings.includes('<anyKey>anyValue</anyKey>'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should set Japanese value as xml string', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let sharedStrings = new Excel(template).setSharedStrings({
                        anyKey: '日本語'
                    }).sharedStrings();
                    assert.isOk(sharedStrings.includes('<anyKey>日本語</anyKey>'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should set value with encoding', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let sharedStrings = new Excel(template).setSharedStrings({
                        anyKey: '<>\"\\\&\''
                    }).sharedStrings();
                    assert.isOk(sharedStrings.includes('<anyKey>\&lt;\&gt;\"\\\&amp;\'</anyKey>'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });
    });

    describe('parseWorkbookRels()', () => {

        it('should parse relation files, styles/sharedStrings/worksheets/theme', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    return new Excel(template).parseWorkbookRels();
                }).then((workbookRels) => {
                    let relationships = workbookRels.Relationships.Relationship;
                    relationships = _.map(relationships, (e) => e['$'].Target);
                    assert.isOk(_.containsAsPartialString(relationships, 'styles.xml'));
                    assert.isOk(_.containsAsPartialString(relationships, 'sharedStrings.xml'));
                    assert.isOk(_.containsAsPartialString(relationships, 'worksheets/'));
                    assert.isOk(_.containsAsPartialString(relationships, 'theme/'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should parse each relation', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template3Sheet.xlsx`)
                .then((template) => {
                    return new Excel(template).parseWorkbookRels();
                }).then((workbookRels) => {
                    let sheetCount = _.chain(workbookRels.Relationships.Relationship)
                        .map((e) => e['$'].Target)
                        .filter((e) => e.includes('worksheets/'))
                        .value()
                        .length;
                    assert.strictEqual(sheetCount, 3);
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });
    });

    describe('setWorkbookRels()', () => {

        it('should return this instance', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let test = new Excel(template).setWorkbookRels({});
                    assert.notStrictEqual(test, undefined);
                    assert.notStrictEqual(test, null);
                    assert.isOk(test instanceof Excel);
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should set value as xml string', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let workbookRels  = new Excel(template)
                        .setWorkbookRels({anyKey: 'anyValue'})
                        .file(config.EXCEL_FILES.FILE_WORKBOOK_RELS)
                        .asText();
                    assert.isOk(workbookRels.includes('<anyKey>anyValue</anyKey>'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should set Japanese value as xml string', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let workbookRels  = new Excel(template)
                        .setWorkbookRels({anyKey: '日本語'})
                        .file(config.EXCEL_FILES.FILE_WORKBOOK_RELS)
                        .asText();
                    assert.isOk(workbookRels.includes('<anyKey>日本語</anyKey>'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should set value with encoding', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let workbookRels  = new Excel(template)
                        .setWorkbookRels({anyKey: '<>\"\\\&\''})
                        .file(config.EXCEL_FILES.FILE_WORKBOOK_RELS)
                        .asText();
                    assert.isOk(workbookRels.includes('<anyKey>\&lt;\&gt;\"\\\&amp;\'</anyKey>'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

    });

    describe('parseWorkbook()', () => {

        it('should parse information of sheet', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    return new Excel(template).parseWorkbook();
                }).then((workbook) => {
                    let sheets = workbook.workbook.sheets[0].sheet;
                    assert.notStrictEqual(sheets, undefined);
                    assert.notStrictEqual(sheets, null);
                    assert.strictEqual(sheets.length, 1);
                    assert.strictEqual(sheets[0]['$'].name, 'Sheet1');
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should parse each sheet', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template3Sheet.xlsx`)
                .then((template) => {
                    return new Excel(template).parseWorkbook();
                }).then((workbook) => {
                    let sheets = workbook.workbook.sheets[0].sheet;
                    assert.notStrictEqual(sheets, undefined);
                    assert.notStrictEqual(sheets, null);
                    assert.strictEqual(sheets.length, 3);
                    assert.strictEqual(sheets[0]['$'].name, 'Sheet1');
                    assert.strictEqual(sheets[1]['$'].name, 'Sheet2');
                    assert.strictEqual(sheets[2]['$'].name, 'Sheet3');
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

    });

    describe('setWorkbook()', () => {

        it('should return this instance', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let test = new Excel(template).setWorkbook({});
                    assert.notStrictEqual(test, undefined);
                    assert.notStrictEqual(test, null);
                    assert.isOk(test instanceof Excel);
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should set value as xml string', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let workbookRels  = new Excel(template)
                        .setWorkbook({anyKey: 'anyValue'})
                        .file(config.EXCEL_FILES.FILE_WORKBOOK)
                        .asText();
                    assert.isOk(workbookRels.includes('<anyKey>anyValue</anyKey>'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should set Japanese value as xml string', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let workbookRels  = new Excel(template)
                        .setWorkbook({anyKey: '日本語'})
                        .file(config.EXCEL_FILES.FILE_WORKBOOK)
                        .asText();
                    assert.isOk(workbookRels.includes('<anyKey>日本語</anyKey>'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should set value with encoding', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let workbookRels  = new Excel(template)
                        .setWorkbook({anyKey: '<>\"\\\&\''})
                        .file(config.EXCEL_FILES.FILE_WORKBOOK)
                        .asText();
                    assert.isOk(workbookRels.includes('<anyKey>\&lt;\&gt;\"\\\&amp;\'</anyKey>'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });
    });

    describe('parseWorksheetsDir()', () => {

        it('should parse relation and contents', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    return new Excel(template).parseWorksheetsDir();
                }).then((worksheets) => {
                    let files = _.filter(worksheets, (e) => !!e.worksheet);
                    assert.strictEqual(files.length, 1);

                    let relations = _.filter(worksheets, (e) => !!e.Relationships);
                    assert.strictEqual(relations.length, 1);

                    assert.strictEqual(`${files[0].name}.rels`, relations[0].name);
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should parse each relation and contents', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template3Sheet.xlsx`)
                .then((template) => {
                    return new Excel(template).parseWorksheetsDir();
                }).then((worksheets) => {
                    let files = _.filter(worksheets, (e) => !!e.worksheet);
                    assert.strictEqual(files.length, 3);

                    let relations = _.filter(worksheets, (e) => !!e.Relationships);
                    assert.strictEqual(relations.length, 3);

                    let fileNameInRelations = _.map(relations, (e) => e.name);
                    assert.isOk(_.contains(fileNameInRelations, `${files[0].name}.rels`));
                    assert.isOk(_.contains(fileNameInRelations, `${files[1].name}.rels`));
                    assert.isOk(_.contains(fileNameInRelations, `${files[2].name}.rels`));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });
    });

    describe('setWorksheet()', () => {

        it('should return this instance', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let test = new Excel(template).setWorksheet('someSheet.xml', {});
                    assert.notStrictEqual(test, undefined);
                    assert.notStrictEqual(test, null);
                    assert.isOk(test instanceof Excel);
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should set value as xml string', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let workSheet = new Excel(template)
                        .setWorksheet('someSheet.xml', {anyKey: 'anyValue'})
                        .file(`${config.EXCEL_FILES.DIR_WORKSHEETS}/someSheet.xml`)
                        .asText();
                    assert.isOk(workSheet.includes('<anyKey>anyValue</anyKey>'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should set Japanese value as xml string', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let workSheet = new Excel(template)
                        .setWorksheet('someSheet.xml', {anyKey: '日本語'})
                        .file(`${config.EXCEL_FILES.DIR_WORKSHEETS}/someSheet.xml`)
                        .asText();
                    assert.isOk(workSheet.includes('<anyKey>日本語</anyKey>'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should set value with encoding', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let workSheet = new Excel(template)
                        .setWorksheet('someSheet.xml', {anyKey: '<>\"\\\&\''})
                        .file(`${config.EXCEL_FILES.DIR_WORKSHEETS}/someSheet.xml`)
                        .asText();
                    assert.isOk(workSheet.includes('<anyKey>\&lt;\&gt;\"\\\&amp;\'</anyKey>'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });
    });

    describe('setWorksheets()', () => {

        it('should return this instance', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let test = new Excel(template).setWorksheets([]);
                    assert.notStrictEqual(test, undefined);
                    assert.notStrictEqual(test, null);
                    assert.isOk(test instanceof Excel);
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should set value as xml string', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let workSheet = new Excel(template)
                        .setWorksheets([
                            {name: 'sheet1.xml', data: {anyKey: 'anyValue'}}
                        ])
                        .file(`${config.EXCEL_FILES.DIR_WORKSHEETS}/sheet1.xml`)
                        .asText();
                    assert.isOk(workSheet.includes('<anyKey>anyValue</anyKey>'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should set each value', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let excelTemplate = new Excel(template)
                        .setWorksheets([
                            {name: 'sheet1.xml', data: {key1: 'value1'}},
                            {name: 'sheet2.xml', data: {key2: 'value2'}},
                            {name: 'sheet3.xml', data: {key3: 'value3'}}
                        ]);
                    let sheet1 = excelTemplate
                        .file(`${config.EXCEL_FILES.DIR_WORKSHEETS}/sheet1.xml`)
                        .asText();
                    assert.isOk(sheet1.includes('<key1>value1</key1>'));

                    let sheet2 = excelTemplate
                        .file(`${config.EXCEL_FILES.DIR_WORKSHEETS}/sheet2.xml`)
                        .asText();
                    assert.isOk(sheet2.includes('<key2>value2</key2>'));

                    let sheet3 = excelTemplate
                        .file(`${config.EXCEL_FILES.DIR_WORKSHEETS}/sheet3.xml`)
                        .asText();
                    assert.isOk(sheet3.includes('<key3>value3</key3>'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should set Japanese value as xml string', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let workSheet = new Excel(template)
                        .setWorksheets([
                            {name: 'sheet1.xml', data: {anyKey: '日本語'}}
                        ])
                        .file(`${config.EXCEL_FILES.DIR_WORKSHEETS}/sheet1.xml`)
                        .asText();
                    assert.isOk(workSheet.includes('<anyKey>日本語</anyKey>'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should set value with encoding', () => {
            return fs.readFileAsync(`${config.TEST_DIRS.TEMPLATE}Template.xlsx`)
                .then((template) => {
                    let workSheet = new Excel(template)
                        .setWorksheets([
                            {name: 'sheet1.xml', data: {anyKey: '<>\"\\\&\''}}
                        ])
                        .file(`${config.EXCEL_FILES.DIR_WORKSHEETS}/sheet1.xml`)
                        .asText();
                    assert.isOk(workSheet.includes('<anyKey>\&lt;\&gt;\"\\\&amp;\'</anyKey>'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });
    });
});