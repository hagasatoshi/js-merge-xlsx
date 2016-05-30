'use strict';

var Promise = require('bluebird');
var _ = require('underscore');
var fs = Promise.promisifyAll(require('fs'));
var WorkBookXml = require('../lib/WorkBookXml');
require('../lib/underscore_mixin');
var assert = require('chai').assert;
var config = require('../lib/Config');
var xml2js = require('xml2js');
var builder = new xml2js.Builder();
var parseString = Promise.promisify(xml2js.parseString);

var readFile = function readFile(xmlFile) {
    return fs.readFileAsync('' + config.TEST_DIRS.XML + xmlFile, 'utf8').then(function (workBookXml) {
        return parseString(workBookXml);
    });
};

describe('WorkBookXml.js', function () {
    describe('constructor', function () {

        it('should set each member from parameter', function () {
            return readFile('workbook.xml').then(function (workBookXml) {
                var workBookXmlObj = new WorkBookXml(workBookXml);
                console.log(workBookXmlObj.workBookXml);
                assert.isOk(_.consistOf(workBookXmlObj, ['workBookXml']));
                //assert.notStrictEqual(workBookXmlObj, undefined);
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });
    });
});