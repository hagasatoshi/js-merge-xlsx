'use strict';

var Promise = require('bluebird');
var _ = require('underscore');
var readYamlAsync = Promise.promisify(require('read-yaml'));
var fs = Promise.promisifyAll(require('fs'));
var excelmerge = require('../ExcelMerge');
var assert = require('chai').assert;

var config = {
    templateDir: './test/templates/',
    testDataDir: './test/data/',
    outptutDir: './test/output/'
};

var removeExstingFiles = function removeExstingFiles(regex) {
    _.each(fs.readdirSync(config.outptutDir), function (file) {
        if (regex && !regex.test(file)) {
            return;
        }
        fs.unlinkSync('' + config.outptutDir + file);
    });
};

var readFiles = function readFiles(template, yaml) {
    return Promise.props({
        template: fs.readFileAsync('' + config.templateDir + template),
        data: readYamlAsync('' + config.testDataDir + yaml)
    });
};

removeExstingFiles(/\.(xlsx|zip)$/);

describe('test for excelmerge.merge()', function () {

    it('excel file is created successfully', function () {
        return readFiles('Template.xlsx', 'data1.yml').then(function (_ref) {
            var template = _ref.template;
            var data = _ref.data;

            return fs.writeFileAsync(config.outptutDir + 'test1.xlsx', excelmerge.merge(template, data));
        }).then(function () {
            assert.isOk(true);
        })['catch'](function (err) {
            console.log(err);
            assert.isOk(false);
        });
    });
});

describe('test for excelmerge.bulkMergeToFiles()', function () {

    it('excel file is created successfully', function () {
        return readFiles('Template.xlsx', 'data2.yml').then(function (_ref2) {
            var template = _ref2.template;
            var data = _ref2.data;

            var arrayObj = _.map(data, function (e, index) {
                return { name: 'file' + index + '.xlsx', data: e };
            });
            return fs.writeFileAsync(config.outptutDir + 'test2.zip', excelmerge.bulkMergeToFiles(template, arrayObj));
        }).then(function () {
            assert.isOk(true);
        })['catch'](function (err) {
            console.log(err);
            assert.isOk(false);
        });
    });
});

describe('test for excelmerge.bulkMergeToSheets()', function () {

    it('excel file is created successfully', function () {
        return readFiles('Template.xlsx', 'data2.yml').then(function (_ref3) {
            var template = _ref3.template;
            var data = _ref3.data;

            var arrayObj = _.map(data, function (e, index) {
                return { name: 'test' + index, data: e };
            });
            return excelmerge.bulkMergeToSheets(template, arrayObj);
        }).then(function (excelData) {
            return fs.writeFileAsync(config.outptutDir + 'test3.xlsx', excelData);
        }).then(function () {
            assert.isOk(true);
        })['catch'](function (err) {
            console.log(err);
            assert.isOk(false);
        });
    });
});