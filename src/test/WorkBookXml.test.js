const Promise = require('bluebird');
const _ = require('underscore');
const fs = Promise.promisifyAll(require('fs'));
const WorkBookXml = require('../lib/WorkBookXml');
require('../lib/underscore_mixin');
const assert = require('chai').assert;
const config = require('../lib/Config');
const xml2js = require('xml2js');
const builder = new xml2js.Builder();
const parseString = Promise.promisify(xml2js.parseString);

const readFile = (xmlFile) => {
    return fs.readFileAsync(`${config.TEST_DIRS.XML}${xmlFile}`, 'utf8')
        .then((workBookXml) => {
            return parseString(workBookXml);
        });
};

describe('WorkBookXml.js', () => {
    describe('constructor', () => {

        it('should set each member from parameter', () => {
            return readFile('workbook.xml')
                .then((workBookXml) => {
                    let workBookXmlObj = new WorkBookXml(workBookXml);
                    assert.isOk(_.consistOf(workBookXmlObj, ['workBookXml', 'sheetDefinitions']));
                    assert.isOk(_.consistOf(workBookXmlObj, ['workBookXml', 'sheetDefinitions']));
                }).catch((err) => {
                    console.log(err);
                    assert.isOk(false);
                });
        });
    });

});