const Promise = require('bluebird');
const _ = require('underscore');
const fs = Promise.promisifyAll(require('fs'));
const WorkBookRels = require('../lib/WorkBookRels');
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

describe('WorkBookRels.js', () => {
    describe('constructor', () => {
        it('should set each member from parameter', () => {
            return readFile('workbook.xml.rels')
            .then((workBookXmlRels) => {
                let workBookXmRelslObj = new WorkBookRels(workBookXmlRels);
                assert.isOk(
                    _.consistOf(workBookXmRelslObj,
                        [
                            {workBookRels:
                                {Relationships: ['$', 'Relationship']}
                            },
                            {sheetRelationships:
                                {$: ['Id', 'Type', 'Target']}
                            }
                        ]
                    )
                );
            });
        });

        it('should have the same number of relationship files', () => {
            return readFile('workbookHaving3Sheet.xml.rels')
            .then((workBookXmlRels) => {
                let workBookXmRelslObj = new WorkBookRels(workBookXmlRels);
                assert.strictEqual(workBookXmRelslObj.sheetRelationships.length, 6);
            });
        });
    });

    describe('add()', () => {
        it('should be added by calling add()', () => {
            return readFile('workbook.xml.rels')
            .then((workBookXmlRels) => {
                let workBookXmRelslObj = new WorkBookRels(workBookXmlRels);
                assert.strictEqual(workBookXmRelslObj.sheetRelationships.length, 4);
                workBookXmRelslObj.add('addedSheet');
                assert.strictEqual(workBookXmRelslObj.sheetRelationships.length, 5);
                let latestRelation = workBookXmRelslObj.sheetRelationships[4];
                assert.strictEqual(latestRelation['$'].Id, 'addedSheet');
                assert.strictEqual(latestRelation['$'].Type, config.OPEN_XML_SCHEMA_DEFINITION);
                assert.strictEqual(latestRelation['$'].Target, 'worksheets/sheetaddedSheet.xml');
            });
        })
    });

    describe('delete()', () => {
        it('should be deleted by calling delete()', () => {
            return readFile('workbook.xml.rels')
            .then((workBookXmlRels) => {
                let workBookXmRelslObj = new WorkBookRels(workBookXmlRels);
                workBookXmRelslObj.add('addedSheet');
                assert.strictEqual(workBookXmRelslObj.sheetRelationships.length, 5);

                workBookXmRelslObj.delete('worksheets/sheetaddedSheet.xml');
                assert.strictEqual(workBookXmRelslObj.sheetRelationships.length, 4);
            });
        });

        it('should not throw error with invalid sheet name', () => {
            return readFile('workbook.xml.rels')
            .then((workBookXmlRels) => {
                let workBookXmRelslObj = new WorkBookRels(workBookXmlRels);
                assert.strictEqual(workBookXmRelslObj.sheetRelationships.length, 4);

                workBookXmRelslObj.delete('worksheets/invalidSheetName.xml');
                assert.strictEqual(workBookXmRelslObj.sheetRelationships.length, 4);
            });
        });

        it('should not throw error by deleting last sheet', () => {
            return readFile('workbookHaving1Element.xml.rels')
            .then((workBookXmlRels) => {
                let workBookXmRelslObj = new WorkBookRels(workBookXmlRels);
                assert.strictEqual(workBookXmRelslObj.sheetRelationships.length, 1);

                workBookXmRelslObj.delete('sharedStrings.xml');
                assert.strictEqual(workBookXmRelslObj.sheetRelationships.length, 0);
            });
        });
    });

    describe('findSheetPath()', () => {

        it('should find by id', () => {
            return readFile('workbook.xml.rels')
            .then((workBookXmlRels) => {
                let workBookXmRelslObj = new WorkBookRels(workBookXmlRels);
                let sheet = workBookXmRelslObj.findSheetPath('rId1');
                assert.strictEqual(sheet, 'worksheets/sheet1.xml');
            });
        });

        it('should return null if invalid id', () => {
            return readFile('workbook.xml.rels')
            .then((workBookXmlRels) => {
                let workBookXmRelslObj = new WorkBookRels(workBookXmlRels);
                let sheet = workBookXmRelslObj.findSheetPath('invalid id');
                assert.strictEqual(sheet, null);
            });
        })
    });

    describe('nextRelationshipId()', () => {
        it('should return next id', () => {
            return readFile('workbook.xml.rels')
            .then((workBookXmlRels) => {
                let nextId = new WorkBookRels(workBookXmlRels).nextRelationshipId();
                assert.strictEqual(nextId, 'rId005');
            });
        });

        it('next id should be sequencial', () => {
            return readFile('workbook.xml.rels')
            .then((workBookXmlRels) => {
                let workBookXmlRelsObj = new WorkBookRels(workBookXmlRels);
                let nextId = workBookXmlRelsObj.nextRelationshipId();
                assert.strictEqual(nextId, 'rId005');
                nextId = workBookXmlRelsObj.add(nextId).nextRelationshipId();
                assert.strictEqual(nextId, 'rId006');
                nextId = workBookXmlRelsObj.add(nextId).nextRelationshipId();
                assert.strictEqual(nextId, 'rId007');
                nextId = workBookXmlRelsObj.add(nextId).nextRelationshipId();
                assert.strictEqual(nextId, 'rId008');
                nextId = workBookXmlRelsObj.add(nextId).nextRelationshipId();
                assert.strictEqual(nextId, 'rId009');
                nextId = workBookXmlRelsObj.add(nextId).nextRelationshipId();
                assert.strictEqual(nextId, 'rId010');
                nextId = workBookXmlRelsObj.add(nextId).nextRelationshipId();
                assert.strictEqual(nextId, 'rId011');
                nextId = workBookXmlRelsObj.add(nextId).nextRelationshipId();
                assert.strictEqual(nextId, 'rId012');
            });
        });

    })
});