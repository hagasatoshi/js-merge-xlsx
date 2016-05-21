const Promise = require('bluebird');
const _ = require('underscore');
const fs = Promise.promisifyAll(require('fs'));
const Excel = require('../lib/Excel');
require('../lib/underscore_mixin');
const assert = require('chai').assert;

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

describe('Excel.js', () => {
    describe('sharedStrings()', () => {

        it('should read each strings on template', () => {
            return fs.readFileAsync(`${config.templateDir}Template.xlsx`)
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
            return fs.readFileAsync(`${config.templateDir}TemplateNoStrings.xlsx`)
                .then((template) => {
                    let sharedStrings = new Excel(template).sharedStrings();
                    assert.isOk(typeof sharedStrings === 'string');
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should read Japanese strings on template', () => {
            return fs.readFileAsync(`${config.templateDir}Template.xlsx`)
                .then((template) => {
                    let sharedStrings = new Excel(template).sharedStrings();
                    assert.isOk(sharedStrings.includes('雇用期間'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should read as encoded string', () => {
            return fs.readFileAsync(`${config.templateDir}TemplateWithXmlEntity.xlsx`)
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
            return fs.readFileAsync(`${config.templateDir}Template.xlsx`)
                .then((template) => {
                    return new Excel(template).parseSharedStrings();
                }).then((templateObj) => {
                    let si = templateObj.sst.si;
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
            return fs.readFileAsync(`${config.templateDir}TemplateNoStrings.xlsx`)
                .then((template) => {
                    return new Excel(template).parseSharedStrings();
                }).then((templateObj) => {
                    assert.isOk(!templateObj.sst || !templateObj.sst.si);
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });

        it('should parse Japanese with no error', () => {
            return fs.readFileAsync(`${config.templateDir}Template.xlsx`)
                .then((template) => {
                    return new Excel(template).parseSharedStrings();
                }).then((templateObj) => {
                    let si = templateObj.sst.si;
                    assert.notStrictEqual(si, undefined);
                    assert.isOk(si instanceof Array);

                    si = _.map(si, (e) => _.stringValue(e.t));
                    assert.isOk(_.containsAsPartialString(si, '雇用期間'));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });
    });
});