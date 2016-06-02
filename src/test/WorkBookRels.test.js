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
    });

});