const Promise = require('bluebird');
const _ = require('underscore');
const fs = Promise.promisifyAll(require('fs'));
const WorkBookXml = require('../lib/WorkBookXml');
require('../lib/underscore');
const assert = require('chai').assert;
const config = require('../lib/Config');
const xml2js = require('xml2js');
const parseString = Promise.promisify(xml2js.parseString);

const readFile = (xmlFile) => {
    return fs.readFileAsync(`${config.TEST_DIRS.XML}${xmlFile}`, 'utf8')
    .then((workBookXml) => {
        return parseString(workBookXml);
    });
};

describe('WorkBookXml.js', () => {
    describe('constructor', () => {
        it('should set each member from parameter', () => {
            return readFile('workbook.xml')
            .then((workBookXml) => {
                let workBookXmlObj = new WorkBookXml(workBookXml);
                assert.isOk(
                    _.consistOf(workBookXmlObj,
                        [
                            {workBookXml:
                                {workbook: [
                                    '$', 'fileVersion', 'workbookPr', 'bookViews',
                                    'sheets', 'definedNames', 'calcPr', 'extLst'
                                ]}
                            },
                            {sheetDefinitions:
                                {$: [
                                    'name', 'sheetId', 'r:id'
                                ]}
                            }
                        ]
                    )
                );
            }).catch((err) => {
                console.log(err);
                assert.isOk(false);
            });
        });
        it('should have the same number of sheet-definitions', () => {
            return readFile('workbookHaving2Sheet.xml')
            .then((workBookXml) => {
                let workBookXmlObj = new WorkBookXml(workBookXml);
                assert.strictEqual(workBookXmlObj.sheetDefinitions.length, 2);
            }).catch((err) => {
                console.log(err);
                assert.isOk(false);
            });
        });
    });

    describe('add()', () => {
        it('should add element formatted as {name, sheetId, r:id}', () => {
            return readFile('workbook.xml')
            .then((workBookXml) => {
                let workBookXmlObj = new WorkBookXml(workBookXml);
                workBookXmlObj.add('addedSheetName', 'addedSheetId');
                assert.strictEqual(workBookXmlObj.sheetDefinitions.length, 2);
                assert.strictEqual(workBookXmlObj.sheetDefinitions[1]['$'].name, 'addedSheetName');
                assert.strictEqual(workBookXmlObj.sheetDefinitions[1]['$'].sheetId, 'addedSheetId');
                assert.strictEqual(workBookXmlObj.sheetDefinitions[1]['$']['r:id'], 'addedSheetId');
            });
        });
    });

    describe('delete()', () => {
        it('should delete correct sheet by name', () => {
            return readFile('workbookHaving2Sheet.xml')
            .then((workBookXml) => {
                let workBookXmlObj = new WorkBookXml(workBookXml);
                workBookXmlObj.delete('Sheet2');
                assert.strictEqual(workBookXmlObj.sheetDefinitions.length, 1);
                assert.strictEqual(workBookXmlObj.sheetDefinitions[0]['$'].name, 'Sheet1');
                assert.strictEqual(workBookXmlObj.sheetDefinitions[0]['$'].sheetId, '1');
                assert.strictEqual(workBookXmlObj.sheetDefinitions[0]['$']['r:id'], 'rId1');
            });
        });

        it('should do nothing with invalid sheet name', () => {
            return readFile('workbookHaving2Sheet.xml')
                .then((workBookXml) => {
                    let workBookXmlObj = new WorkBookXml(workBookXml);
                    workBookXmlObj.delete('invalid-sheet-name');
                    assert.strictEqual(workBookXmlObj.sheetDefinitions.length, 2);
                    assert.strictEqual(workBookXmlObj.sheetDefinitions[0]['$'].name, 'Sheet1');
                    assert.strictEqual(workBookXmlObj.sheetDefinitions[0]['$'].sheetId, '1');
                    assert.strictEqual(workBookXmlObj.sheetDefinitions[0]['$']['r:id'], 'rId1');
                    assert.strictEqual(workBookXmlObj.sheetDefinitions[1]['$'].name, 'Sheet2');
                    assert.strictEqual(workBookXmlObj.sheetDefinitions[1]['$'].sheetId, '2');
                    assert.strictEqual(workBookXmlObj.sheetDefinitions[1]['$']['r:id'], 'rId2');
                });
        });
    });

    describe('findSheetId()', () => {
        it('should return null if invalid sheet name', () => {
            return readFile('workbookHaving2Sheet.xml')
            .then((workBookXml) => {
                let sheetId = new WorkBookXml(workBookXml).findSheetId('invalid-sheet-name');
                assert.strictEqual(sheetId, null);
            });
        });

        it('should return correct sheet data by name', () => {
            return readFile('workbookHaving2Sheet.xml')
            .then((workBookXml) => {
                let sheetId = new WorkBookXml(workBookXml).findSheetId('Sheet1');
                assert.notStrictEqual(sheetId, null);
                assert.strictEqual(sheetId, 'rId1');
            });
        });
    });

    describe('firstSheetName()', () => {
        it('should return name of the first sheet', () => {
            return readFile('workbookHaving2Sheet.xml')
            .then((workBookXml) => {
                let firstSheetName = new WorkBookXml(workBookXml).firstSheetName();
                assert.strictEqual(firstSheetName, 'Sheet1');
            });
        });
    });

    describe('value()', () => {
        it('should retrieve the latest value', () => {
            return readFile('workbook.xml')
            .then((workBookXml) => {
                let sheets = new WorkBookXml(workBookXml)
                    .add('addedSheetName', 'addedSheetId')
                    .value()
                    .workbook.sheets[0].sheet;
                assert.strictEqual(sheets.length, 2);
                assert.strictEqual(sheets[0]['$'].name, 'Sheet1');
                assert.strictEqual(sheets[0]['$'].sheetId, '1');
                assert.strictEqual(sheets[0]['$']['r:id'], 'rId1');
                assert.strictEqual(sheets[1]['$'].name, 'addedSheetName');
                assert.strictEqual(sheets[1]['$'].sheetId, 'addedSheetId');
                assert.strictEqual(sheets[1]['$']['r:id'], 'addedSheetId');
            });
        });
    });

});