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

describe('Excel.prototype.variables', () => {

    it('should read each variables on template', () => {
        return fs.readFileAsync(`${config.templateDir}Template.xlsx`)
            .then((template) => {
                let variables = new Excel(template).variables();
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
            }).catch((err) => {
                console.log(err);
                assert.ok(false);
            });
    });

    it('should read as empty array in case no strings defined', () => {
        return fs.readFileAsync(`${config.templateDir}TemplateNoStrings.xlsx`)
            .then((template) => {
                let variables = new Excel(template).variables();
                assert.notEqual(undefined, variables);
                assert.notEqual(null, variables);
                assert.ok(variables instanceof Array);
                assert.equal(0, variables.length);
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

describe('Excel.prototype.hasAsSharedString', () => {

    it('should return true if defined', () => {
        return fs.readFileAsync(`${config.templateDir}Template.xlsx`)
            .then((template) => {
                let templateObj = new Excel(template);
                assert.equal(true, templateObj.hasAsSharedString('AccountName__c'));
                assert.equal(true, templateObj.hasAsSharedString('AccountAddress__c'));
                assert.equal(true, templateObj.hasAsSharedString('SalaryDate__c'));
            }).catch((err) => {
                console.log(err);
                assert.ok(false);
            });
    });

    it('should return false if not defined', () => {
        return fs.readFileAsync(`${config.templateDir}Template.xlsx`)
            .then((template) => {
                let templateObj = new Excel(template);
                assert.equal(false, templateObj.hasAsSharedString('invalid string'));
            }).catch((err) => {
                console.log(err);
                assert.ok(false);
            });
    });

    it('should return false with no error in case no strings defined', () => {
        return fs.readFileAsync(`${config.templateDir}TemplateNoStrings.xlsx`)
            .then((template) => {
                let templateObj = new Excel(template);
                assert.equal(false, templateObj.hasAsSharedString('any string'));
            }).catch((err) => {
                console.log(err);
                assert.ok(false);
            });
    });

    it('should return true even if Japanese', () => {
        return fs.readFileAsync(`${config.templateDir}Template.xlsx`)
            .then((template) => {
                let templateObj = new Excel(template);
                assert.equal(true, templateObj.hasAsSharedString('雇用期間'));
            }).catch((err) => {
                console.log(err);
                assert.ok(false);
            });
    });

    it('should match as encoded strings', () => {
        return fs.readFileAsync(`${config.templateDir}TemplateWithXmlEntity.xlsx`)
            .then((template) => {
                let templateObj = new Excel(template);
                assert.equal(true, templateObj.hasAsSharedString('\&lt;\&gt;\"\\\&amp;\''));
                assert.equal(false, templateObj.hasAsSharedString('<>\"\\\&\''));
            }).catch((err) => {
                console.log(err);
                assert.ok(false);
            });
    });

});