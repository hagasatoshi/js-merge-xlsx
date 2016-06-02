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
    });
});