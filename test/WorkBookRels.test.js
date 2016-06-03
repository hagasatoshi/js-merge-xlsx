'use strict';

var Promise = require('bluebird');
var _ = require('underscore');
var fs = Promise.promisifyAll(require('fs'));
var WorkBookRels = require('../lib/WorkBookRels');
require('../lib/underscore');
var assert = require('chai').assert;
var config = require('../lib/Config');
var xml2js = require('xml2js');
var parseString = Promise.promisify(xml2js.parseString);

var readFile = function readFile(xmlFile) {
    return fs.readFileAsync('' + config.TEST_DIRS.XML + xmlFile, 'utf8').then(function (workBookXml) {
        return parseString(workBookXml);
    });
};

describe('WorkBookRels.js', function () {
    describe('constructor', function () {
        it('should set each member from parameter', function () {
            return readFile('workbook.xml.rels').then(function (workBookXmlRels) {
                var workBookXmRelslObj = new WorkBookRels(workBookXmlRels);
                assert.isOk(_.consistOf(workBookXmRelslObj, [{ workBookRels: { Relationships: ['$', 'Relationship'] }
                }, { sheetRelationships: { $: ['Id', 'Type', 'Target'] }
                }]));
            });
        });

        it('should have the same number of relationship files', function () {
            return readFile('workbookHaving3Sheet.xml.rels').then(function (workBookXmlRels) {
                var workBookXmRelslObj = new WorkBookRels(workBookXmlRels);
                assert.strictEqual(workBookXmRelslObj.sheetRelationships.length, 6);
            });
        });
    });

    describe('add()', function () {
        it('should be added by calling add()', function () {
            return readFile('workbook.xml.rels').then(function (workBookXmlRels) {
                var workBookXmRelslObj = new WorkBookRels(workBookXmlRels);
                assert.strictEqual(workBookXmRelslObj.sheetRelationships.length, 4);
                workBookXmRelslObj.add('addedSheet');
                assert.strictEqual(workBookXmRelslObj.sheetRelationships.length, 5);
                var latestRelation = workBookXmRelslObj.sheetRelationships[4];
                assert.strictEqual(latestRelation['$'].Id, 'addedSheet');
                assert.strictEqual(latestRelation['$'].Type, config.OPEN_XML_SCHEMA_DEFINITION);
                assert.strictEqual(latestRelation['$'].Target, 'worksheets/sheetaddedSheet.xml');
            });
        });
    });

    describe('delete()', function () {
        it('should be deleted by calling delete()', function () {
            return readFile('workbook.xml.rels').then(function (workBookXmlRels) {
                var workBookXmRelslObj = new WorkBookRels(workBookXmlRels);
                workBookXmRelslObj.add('addedSheet');
                assert.strictEqual(workBookXmRelslObj.sheetRelationships.length, 5);

                workBookXmRelslObj['delete']('worksheets/sheetaddedSheet.xml');
                assert.strictEqual(workBookXmRelslObj.sheetRelationships.length, 4);
            });
        });

        it('should not throw error with invalid sheet name', function () {
            return readFile('workbook.xml.rels').then(function (workBookXmlRels) {
                var workBookXmRelslObj = new WorkBookRels(workBookXmlRels);
                assert.strictEqual(workBookXmRelslObj.sheetRelationships.length, 4);

                workBookXmRelslObj['delete']('worksheets/invalidSheetName.xml');
                assert.strictEqual(workBookXmRelslObj.sheetRelationships.length, 4);
            });
        });

        it('should not throw error by deleting last sheet', function () {
            return readFile('workbookHaving1Element.xml.rels').then(function (workBookXmlRels) {
                var workBookXmRelslObj = new WorkBookRels(workBookXmlRels);
                assert.strictEqual(workBookXmRelslObj.sheetRelationships.length, 1);

                workBookXmRelslObj['delete']('sharedStrings.xml');
                assert.strictEqual(workBookXmRelslObj.sheetRelationships.length, 0);
            });
        });
    });

    describe('findSheetPath()', function () {

        it('should find by id', function () {
            return readFile('workbook.xml.rels').then(function (workBookXmlRels) {
                var workBookXmRelslObj = new WorkBookRels(workBookXmlRels);
                var sheet = workBookXmRelslObj.findSheetPath('rId1');
                assert.strictEqual(sheet, 'worksheets/sheet1.xml');
            });
        });

        it('should return null if invalid id', function () {
            return readFile('workbook.xml.rels').then(function (workBookXmlRels) {
                var workBookXmRelslObj = new WorkBookRels(workBookXmlRels);
                var sheet = workBookXmRelslObj.findSheetPath('invalid id');
                assert.strictEqual(sheet, null);
            });
        });
    });

    describe('nextRelationshipId()', function () {
        it('should return next id', function () {
            return readFile('workbook.xml.rels').then(function (workBookXmlRels) {
                var nextId = new WorkBookRels(workBookXmlRels).nextRelationshipId();
                assert.strictEqual(nextId, 'rId005');
            });
        });

        it('next id should be sequencial', function () {
            return readFile('workbook.xml.rels').then(function (workBookXmlRels) {
                var workBookXmlRelsObj = new WorkBookRels(workBookXmlRels);
                var nextId = workBookXmlRelsObj.nextRelationshipId();
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
    });
});