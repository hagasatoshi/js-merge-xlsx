'use strict';

var Promise = require('bluebird');
var _ = require('underscore');
var fs = Promise.promisifyAll(require('fs'));
var Excel = require('../lib/Excel');
require('../lib/underscore_mixin');
var assert = require('chai').assert;

var config = {
    templateDir: './test/templates/',
    testDataDir: './test/data/',
    outptutDir: './test/output/'
};

var readFiles = function readFiles(template) {
    return Promise.props({
        template: fs.readFileAsync('' + config.templateDir + template)
    });
};

describe('Excel.js', function () {
    describe('sharedStrings()', function () {

        it('should read each strings on template', function () {
            return fs.readFileAsync(config.templateDir + 'Template.xlsx').then(function (template) {
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
            return fs.readFileAsync(config.templateDir + 'TemplateNoStrings.xlsx').then(function (template) {
                var sharedStrings = new Excel(template).sharedStrings();
                assert.isOk(typeof sharedStrings === 'string');
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should read Japanese strings on template', function () {
            return fs.readFileAsync(config.templateDir + 'Template.xlsx').then(function (template) {
                var sharedStrings = new Excel(template).sharedStrings();
                assert.isOk(sharedStrings.includes('雇用期間'));
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should read as encoded string', function () {
            return fs.readFileAsync(config.templateDir + 'TemplateWithXmlEntity.xlsx').then(function (template) {
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
            return fs.readFileAsync(config.templateDir + 'Template.xlsx').then(function (template) {
                return new Excel(template).parseSharedStrings();
            }).then(function (templateObj) {
                var si = templateObj.sst.si;
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
            return fs.readFileAsync(config.templateDir + 'TemplateNoStrings.xlsx').then(function (template) {
                return new Excel(template).parseSharedStrings();
            }).then(function (templateObj) {
                assert.isOk(!templateObj.sst || !templateObj.sst.si);
            })['catch'](function (err) {
                console.log(err);
                assert.isOk(false);
            });
        });

        it('should parse Japanese with no error', function () {
            return fs.readFileAsync(config.templateDir + 'Template.xlsx').then(function (template) {
                return new Excel(template).parseSharedStrings();
            }).then(function (templateObj) {
                var si = templateObj.sst.si;
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
    });
});