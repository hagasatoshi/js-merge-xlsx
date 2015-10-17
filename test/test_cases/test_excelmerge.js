/**
 * * test_spreadsheet.js
 * * Test code for spreadsheet
 * * @author Satoshi Haga
 * * @date 2015/10/10
 **/
'use strict';

var path = require('path');
var cwd = path.resolve('');
var assert = require('assert');
var JSZip = require('jszip');
var ExcelMerge = require(cwd + '/excelmerge');
var SpreadSheet = require(cwd + '/lib/spreadsheet');
require(cwd + '/lib/underscore_mixin');
var Promise = require('bluebird');
var readYamlAsync = Promise.promisify(require('read-yaml'));
var fs = Promise.promisifyAll(require('fs'));
var _ = require('underscore');

module.exports = {
    checkLoadWithNoParameterShouldReturnError: function checkLoadWithNoParameterShouldReturnError() {
        return new SpreadSheet().load().then(function () {
            throw new Error('test_load_with_no_parameter_should_return_error failed ');
        })['catch'](function (err) {
            assert.equal(err, 'First parameter must be JSZip instance including MS-Excel data');
        });
    },

    checkLoadShouldReturnThisInstance: function checkLoadShouldReturnThisInstance() {
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (validTemplate) {
            return new ExcelMerge().load(new JSZip(validTemplate));
        }).then(function (excelMerge) {
            assert(excelMerge instanceof ExcelMerge, 'ExcelMerge#load() should return this instance');
        });
    },

    checkIfBulkMergeMultiSheetRendersCorrectly: function checkIfBulkMergeMultiSheetRendersCorrectly() {
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (validTemplate) {
            return new ExcelMerge().load(new JSZip(validTemplate));
        }).then(function (excelMerge) {
            return excelMerge.bulkMergeMultiSheet([{ name: 'sheet1', data: { AccountName__c: 'hoge account1', AccountAddress__c: 'hoge street1' } }, { name: 'sheet2', data: { AccountName__c: 'hoge account2', AccountAddress__c: 'hoge street2' } }, { name: 'sheet3', data: { AccountName__c: 'hoge account3', AccountAddress__c: 'hoge street3' } }]);
        }).then(function (excelData) {
            return new SpreadSheet().load(new JSZip(excelData));
        }).then(function (spreadsheet) {
            assert(spreadsheet.hasAsSharedString('hoge account1'), "'hoge account1' is missing in excel file");
            assert(spreadsheet.hasAsSharedString('hoge street1'), "'hoge street1' is missing in excel file");
            assert(spreadsheet.hasAsSharedString('hoge account2'), "'hoge account2' is missing in excel file");
            assert(spreadsheet.hasAsSharedString('hoge street2'), "'hoge street2' is missing in excel file");
            assert(spreadsheet.hasAsSharedString('hoge account3'), "'hoge account3' is missing in excel file");
            assert(spreadsheet.hasAsSharedString('hoge street3'), "'hoge street3' is missing in excel file");
        });
    },

    checkLoadEachMemberFromValidTemplate: function checkLoadEachMemberFromValidTemplate() {
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (validTemplate) {
            return new ExcelMerge().load(new JSZip(validTemplate));
        }).then(function (excelMerge) {

            //excel
            assert(excelMerge.spreadsheet.excel instanceof JSZip, 'SpreadSheet#excel is not assigned correctly');

            //check if each variables is parsed or not.
            var variables = ['AccountName__c', 'StartDateFormat__c', 'EndDateFormat__c', 'JobDescription__c', 'StartTime__c', 'EndTime__c', 'hasOverTime__c', 'HoliDayType__c', 'Salary__c', 'DueDate__c', 'SalaryDate__c', 'AccountName__c', 'AccountAddress__c'];
            var chkCommonStringsWithVariable = _.map(excelMerge.spreadsheet.commonStringsWithVariable, function (e) {
                return _(e.t).stringValue();
            });
            _.each(variables, function (e) {
                //variables
                assert(_.contains(excelMerge.spreadsheet.variables, e), 'ExcelMerge#load() doesn\'t set up ' + e + ' as variable correctly');
                assert(_.find(chkCommonStringsWithVariable, function (v) {
                    return v.indexOf('{{' + e + '}}') !== -1;
                }), 'ExcelMerge#load() doesn\'t set up ' + e + ' as variable correctly');
            });
        });
    },

    checkIfMergeRendersCorrectly: function checkIfMergeRendersCorrectly() {
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (validTemplate) {
            return new ExcelMerge().load(new JSZip(validTemplate));
        }).then(function (excelMerge) {
            return excelMerge.merge({ AccountName__c: 'hoge account', AccountAddress__c: 'hoge street' });
        }).then(function (excelData) {
            return new SpreadSheet().load(new JSZip(excelData));
        }).then(function (spreadsheet) {
            assert(spreadsheet.variables.length === 0, "ExcelMerge#merge() doesn't work correctly");
            assert(spreadsheet.hasAsSharedString('hoge account'), "'hoge account' is not rendered by SpreadSheet#simpleMerge()");
            assert(spreadsheet.hasAsSharedString('hoge street'), "'hoge street' is not rendered by SpreadSheet#simpleMerge()");
        });
    },

    checkIfMergeWithNoParameterRendersCorrectly: function checkIfMergeWithNoParameterRendersCorrectly() {
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (validTemplate) {
            return new ExcelMerge().load(new JSZip(validTemplate));
        }).then(function (excelMerge) {
            return excelMerge.merge();
        }).then(function () {
            throw new Error('checkIfMergeWithNoParameterRendersCorrectly failed');
        })['catch'](function (err) {
            assert.equal(err.message, 'merge() must has parameter');
        });
    },

    checkIfBulkMergeMultiFileRendersCorrectly: function checkIfBulkMergeMultiFileRendersCorrectly() {
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (validTemplate) {
            return new ExcelMerge().load(new JSZip(validTemplate));
        }).then(function (excelMerge) {
            return excelMerge.bulkMergeMultiFile([{ name: 'file1.xlsx', data: { AccountName__c: 'hoge account1', AccountAddress__c: 'hoge street1' } }, { name: 'file2.xlsx', data: { AccountName__c: 'hoge account2', AccountAddress__c: 'hoge street2' } }, { name: 'file3.xlsx', data: { AccountName__c: 'hoge account3', AccountAddress__c: 'hoge street3' } }]);
        }).then(function (zipData) {
            var zip = new JSZip(zipData);
            var excel1 = zip.file('file1.xlsx').asArrayBuffer();
            var excel2 = zip.file('file2.xlsx').asArrayBuffer();
            var excel3 = zip.file('file3.xlsx').asArrayBuffer();
            return Promise.props({
                sp1: new SpreadSheet().load(new JSZip(excel1)),
                sp2: new SpreadSheet().load(new JSZip(excel2)),
                sp3: new SpreadSheet().load(new JSZip(excel3))
            }).then(function (_ref) {
                var sp1 = _ref.sp1;
                var sp2 = _ref.sp2;
                var sp3 = _ref.sp3;

                assert(sp1.hasAsSharedString('hoge account1'), "'hoge account1' is missing in excel file");
                assert(sp1.hasAsSharedString('hoge street1'), "'hoge street1' is missing in excel file");
                assert(sp2.hasAsSharedString('hoge account2'), "'hoge account2' is missing in excel file");
                assert(sp2.hasAsSharedString('hoge street2'), "'hoge street2' is missing in excel file");
                assert(sp3.hasAsSharedString('hoge account3'), "'hoge account3' is missing in excel file");
                assert(sp3.hasAsSharedString('hoge street3'), "'hoge street3' is missing in excel file");
            });
        });
    },

    checkIfBulkMergeMultiFileWithNoParameterShouldReturnError: function checkIfBulkMergeMultiFileWithNoParameterShouldReturnError() {
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (validTemplate) {
            return new ExcelMerge().load(new JSZip(validTemplate));
        }).then(function (excelMerge) {
            return excelMerge.bulkMergeMultiFile();
        }).then(function () {
            throw new Error('checkIfBulkMergeMultiFileWithNoParameterShouldReturnError failed');
        })['catch'](function (err) {
            assert.equal(err.message, 'bulkMergeMultiFile() must has parameter');
        });
    },

    checkIfBulkMergeMultiSheetWithNoParameterShouldReturnError: function checkIfBulkMergeMultiSheetWithNoParameterShouldReturnError() {
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (validTemplate) {
            return new ExcelMerge().load(new JSZip(validTemplate));
        }).then(function (excelMerge) {
            return excelMerge.bulkMergeMultiSheet();
        }).then(function () {
            throw new Error('checkIfBulkMergeMultiSheetWithNoParameterShouldReturnError failed');
        })['catch'](function (err) {
            assert.equal(err.message, 'bulkMergeMultiSheet() must has array as parameter');
        });
    }

};