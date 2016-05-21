const Promise = require('bluebird');
const _ = require('underscore');
const fs = Promise.promisifyAll(require('fs'));
const Excel = require('../lib/Excel');
require('../lib/underscore_mixin');
const assert = require('assert');

const config = {
    templateDir: './test/templates/',
    testDataDir: './test/data/',
    outptutDir:  './test/output/'
};

const readFiles = (template) => {
    return Promise.props({
        template: fs.readFileAsync(`${config.templateDir}${template}`)
    });
};

describe('Excel.prototype.sharedStrings', () => {

    it('should read each strings on template', () => {
        return fs.readFileAsync(`${config.templateDir}Template.xlsx`)
            .then((template) => {
                let sharedStrings = new Excel(template).sharedStrings();
                assert.ok(typeof sharedStrings === 'string');
                assert.ok(sharedStrings.includes('{{AccountName__c}}'));
                assert.ok(sharedStrings.includes('{{AccountAddress__c}}'));
                assert.ok(sharedStrings.includes('{{SalaryDate__c}}'));
            }).catch((err) => {
                console.log(err);
                assert.ok(false);
            });
    });

    it('should read with no error in case no strings defined', () => {
        return fs.readFileAsync(`${config.templateDir}TemplateNoStrings.xlsx`)
            .then((template) => {
                let sharedStrings = new Excel(template).sharedStrings();
                assert.ok(typeof sharedStrings === 'string');
            }).catch((err) => {
                console.log(err);
                assert.ok(false);
            });
    });

    it('should read Japanese strings on template', () => {
        return fs.readFileAsync(`${config.templateDir}Template.xlsx`)
            .then((template) => {
                let sharedStrings = new Excel(template).sharedStrings();
                assert.ok(sharedStrings.includes('雇用期間'));
            }).catch((err) => {
                console.log(err);
                assert.ok(false);
            });
    });

    it('should read as encoded string', () => {
        return fs.readFileAsync(`${config.templateDir}TemplateWithXmlEntity.xlsx`)
            .then((template) => {
                let sharedStrings = new Excel(template).sharedStrings();
                assert.ok(sharedStrings.includes('\&lt;\&gt;\"\\\&amp;\''));
            }).catch((err) => {
                console.log(err);
                assert.ok(false);
            });
    });

});

describe('Excel.prototype.parseSharedStrings', () => {

    it('should parse each strings on template', () => {
        return fs.readFileAsync(`${config.templateDir}Template.xlsx`)
            .then((template) => {
                return new Excel(template).parseSharedStrings();
            }).then((templateObj) => {
                let si = templateObj.sst.si;
                assert.notEqual(undefined, si);
                assert.ok(si instanceof Array);

                si = _.map(si, (e) => _.stringValue(e.t));
                assert.ok(_.containsAsPartialString(si, '{{AccountName__c}}'));
                assert.ok(_.containsAsPartialString(si, '{{AccountAddress__c}}'));
                assert.ok(_.containsAsPartialString(si, '{{SalaryDate__c}}'));
            }).catch((err) => {
                console.log(err);
                assert.ok(false);
            });
    });

    it('should parse with no error in case no strings defined', () => {
        return fs.readFileAsync(`${config.templateDir}TemplateNoStrings.xlsx`)
            .then((template) => {
                return new Excel(template).parseSharedStrings();
            }).then((templateObj) => {
                assert.ok(!templateObj.sst || !templateObj.sst.si);
            }).catch((err) => {
                console.log(err);
                assert.ok(false);
            });
    });

    it('should parse Japanese with no error', () => {
        return fs.readFileAsync(`${config.templateDir}Template.xlsx`)
            .then((template) => {
                return new Excel(template).parseSharedStrings();
            }).then((templateObj) => {
                let si = templateObj.sst.si;
                assert.notEqual(undefined, si);
                assert.ok(si instanceof Array);

                si = _.map(si, (e) => _.stringValue(e.t));
                assert.ok(_.containsAsPartialString(si, '雇用期間'));
            }).catch((err) => {
                console.log(err);
                assert.ok(false);
            });
    });
});
