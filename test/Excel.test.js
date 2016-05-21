'use strict';

var Promise = require('bluebird');
var _ = require('underscore');
var fs = Promise.promisifyAll(require('fs'));
var Excel = require('../lib/Excel');
require('../lib/underscore_mixin');
var assert = require('assert');

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

describe('Excel.prototype.sharedStrings', function () {

    it('should read each strings on template', function () {
        return fs.readFileAsync(config.templateDir + 'Template.xlsx').then(function (template) {
            var sharedStrings = new Excel(template).sharedStrings();
            assert.ok(typeof sharedStrings === 'string');
            assert.ok(sharedStrings.includes('{{AccountName__c}}'));
            assert.ok(sharedStrings.includes('{{AccountAddress__c}}'));
            assert.ok(sharedStrings.includes('{{SalaryDate__c}}'));
        })['catch'](function (err) {
            console.log(err);
            assert.ok(false);
        });
    });

    it('should read with no error in case no strings defined', function () {
        return fs.readFileAsync(config.templateDir + 'TemplateNoStrings.xlsx').then(function (template) {
            var sharedStrings = new Excel(template).sharedStrings();
            assert.ok(typeof sharedStrings === 'string');
        })['catch'](function (err) {
            console.log(err);
            assert.ok(false);
        });
    });

    it('should read Japanese strings on template', function () {
        return fs.readFileAsync(config.templateDir + 'Template.xlsx').then(function (template) {
            var sharedStrings = new Excel(template).sharedStrings();
            assert.ok(sharedStrings.includes('雇用期間'));
        })['catch'](function (err) {
            console.log(err);
            assert.ok(false);
        });
    });

    it('should read as encoded string', function () {
        return fs.readFileAsync(config.templateDir + 'TemplateWithXmlEntity.xlsx').then(function (template) {
            var sharedStrings = new Excel(template).sharedStrings();
            assert.ok(sharedStrings.includes('\&lt;\&gt;\"\\\&amp;\''));
        })['catch'](function (err) {
            console.log(err);
            assert.ok(false);
        });
    });
});

describe('Excel.prototype.variables', function () {

    it('should read each variables on template', function () {
        return fs.readFileAsync(config.templateDir + 'Template.xlsx').then(function (template) {
            var variables = new Excel(template).variables();
            assert.ok(variables instanceof Array);
            assert.equal(14, variables.length);
            assert.equal('AccountName__c', variables[0]);
            assert.equal('StartDateFormat__c', variables[1]);
            assert.equal('EndDateFormat__c', variables[2]);
            assert.equal('Address__c', variables[3]);
            assert.equal('JobDescription__c', variables[4]);
            assert.equal('StartTime__c', variables[5]);
            assert.equal('EndTime__c', variables[6]);
            assert.equal('hasOverTime__c', variables[7]);
            assert.equal('HoliDayType__c', variables[8]);
            assert.equal('Salary__c', variables[9]);
            assert.equal('DueDate__c', variables[10]);
            assert.equal('SalaryDate__c', variables[11]);
            assert.equal('AccountName__c', variables[12]);
            assert.equal('AccountAddress__c', variables[13]);
        })['catch'](function (err) {
            console.log(err);
            assert.ok(false);
        });
    });

    it('should read as empty array in case no strings defined', function () {
        return fs.readFileAsync(config.templateDir + 'TemplateNoStrings.xlsx').then(function (template) {
            var variables = new Excel(template).variables();
            assert.notEqual(undefined, variables);
            assert.notEqual(null, variables);
            assert.ok(variables instanceof Array);
            assert.equal(0, variables.length);
        })['catch'](function (err) {
            console.log(err);
            assert.ok(false);
        });
    });
});

describe('Excel.prototype.parseSharedStrings', function () {

    it('should parse each strings on template', function () {
        return fs.readFileAsync(config.templateDir + 'Template.xlsx').then(function (template) {
            return new Excel(template).parseSharedStrings();
        }).then(function (templateObj) {
            var si = templateObj.sst.si;
            assert.notEqual(undefined, si);
            assert.ok(si instanceof Array);

            si = _.map(si, function (e) {
                return _.stringValue(e.t);
            });
            assert.ok(_.containsAsPartialString(si, '{{AccountName__c}}'));
            assert.ok(_.containsAsPartialString(si, '{{AccountAddress__c}}'));
            assert.ok(_.containsAsPartialString(si, '{{SalaryDate__c}}'));
        })['catch'](function (err) {
            console.log(err);
            assert.ok(false);
        });
    });

    it('should parse with no error in case no strings defined', function () {
        return fs.readFileAsync(config.templateDir + 'TemplateNoStrings.xlsx').then(function (template) {
            return new Excel(template).parseSharedStrings();
        }).then(function (templateObj) {
            assert.ok(!templateObj.sst || !templateObj.sst.si);
        })['catch'](function (err) {
            console.log(err);
            assert.ok(false);
        });
    });

    it('should parse Japanese with no error', function () {
        return fs.readFileAsync(config.templateDir + 'Template.xlsx').then(function (template) {
            return new Excel(template).parseSharedStrings();
        }).then(function (templateObj) {
            var si = templateObj.sst.si;
            assert.notEqual(undefined, si);
            assert.ok(si instanceof Array);

            si = _.map(si, function (e) {
                return _.stringValue(e.t);
            });
            assert.ok(_.containsAsPartialString(si, '雇用期間'));
        })['catch'](function (err) {
            console.log(err);
            assert.ok(false);
        });
    });
});

describe('Excel.prototype.hasAsSharedString', function () {

    it('should return true if defined', function () {
        return fs.readFileAsync(config.templateDir + 'Template.xlsx').then(function (template) {
            var templateObj = new Excel(template);
            assert.equal(true, templateObj.hasAsSharedString('AccountName__c'));
            assert.equal(true, templateObj.hasAsSharedString('AccountAddress__c'));
            assert.equal(true, templateObj.hasAsSharedString('SalaryDate__c'));
        })['catch'](function (err) {
            console.log(err);
            assert.ok(false);
        });
    });

    it('should return false if not defined', function () {
        return fs.readFileAsync(config.templateDir + 'Template.xlsx').then(function (template) {
            var templateObj = new Excel(template);
            assert.equal(false, templateObj.hasAsSharedString('invalid string'));
        })['catch'](function (err) {
            console.log(err);
            assert.ok(false);
        });
    });

    it('should return false with no error in case no strings defined', function () {
        return fs.readFileAsync(config.templateDir + 'TemplateNoStrings.xlsx').then(function (template) {
            var templateObj = new Excel(template);
            assert.equal(false, templateObj.hasAsSharedString('any string'));
        })['catch'](function (err) {
            console.log(err);
            assert.ok(false);
        });
    });

    it('should return true even if Japanese', function () {
        return fs.readFileAsync(config.templateDir + 'Template.xlsx').then(function (template) {
            var templateObj = new Excel(template);
            assert.equal(true, templateObj.hasAsSharedString('雇用期間'));
        })['catch'](function (err) {
            console.log(err);
            assert.ok(false);
        });
    });

    it('should match as encoded strings', function () {
        return fs.readFileAsync(config.templateDir + 'TemplateWithXmlEntity.xlsx').then(function (template) {
            var templateObj = new Excel(template);
            assert.equal(true, templateObj.hasAsSharedString('\&lt;\&gt;\"\\\&amp;\''));
            assert.equal(false, templateObj.hasAsSharedString('<>\"\\\&\''));
        })['catch'](function (err) {
            console.log(err);
            assert.ok(false);
        });
    });
});