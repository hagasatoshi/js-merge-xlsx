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
var SpreadSheet = require(cwd + '/lib/spreadsheet');
require(cwd + '/lib/underscore_mixin');
var Promise = require('bluebird');
var readYamlAsync = Promise.promisify(require('read-yaml'));
var fs = Promise.promisifyAll(require('fs'));
var _ = require('underscore');

var test_cases = {};

test_cases.check_load_with_no_parameter_should_return_error = function () {
    var spreadsheet = new SpreadSheet();
    return spreadsheet.load().then(function () {
        throw new Error('test_load_with_no_parameter_should_return_error failed ');
    })['catch'](function (err) {
        assert(err === 'First parameter must be JSZip instance including MS-Excel data');
    });
};

test_cases.check_load_should_return_this_instance = function () {
    var spreadsheet = new SpreadSheet();
    return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (valid_template) {
        return spreadsheet.load(new JSZip(valid_template));
    }).then(function (spreadsheet) {
        assert(spreadsheet instanceof SpreadSheet);
    });
};

test_cases.check_load_with_invalid_sheetname_should_return_error = function () {
    var spreadsheet = new SpreadSheet();
    return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (valid_template) {
        return spreadsheet.load(new JSZip(valid_template), { sheetname: 'hogehoge' });
    }).then(function (spreadsheet) {
        throw new Error('check_load_with_invalid_sheetname_should_throw_error failed ');
    })['catch'](function (err) {
        assert(err === "sheetname is invalid. Please check if sheet'hogehoge' exists in tempalte file");
    });
};

test_cases.check_load_each_member_from_valid_template = function () {
    var spreadsheet = new SpreadSheet();
    return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (valid_template) {
        return spreadsheet.load(new JSZip(valid_template));
    }).then(function (spreadsheet) {

        //excel
        assert(spreadsheet.excel instanceof JSZip);
        //check if each variables is parsed or not.
        var variables = ['AccountName__c', 'StartDateFormat__c', 'EndDateFormat__c', 'JobDescription__c', 'StartTime__c', 'EndTime__c', 'hasOverTime__c', 'HoliDayType__c', 'Salary__c', 'DueDate__c', 'SalaryDate__c', 'AccountName__c', 'AccountAddress__c'];
        var chk_common_strings_with_variable = _.map(spreadsheet.common_strings_with_variable, function (e) {
            return _(e.t).string_value();
        });
        _.each(variables, function (e) {
            //variables
            assert(_.contains(spreadsheet.variables, e));
            assert(_.find(chk_common_strings_with_variable, function (v) {
                return v.indexOf('{{' + e + '}}') !== -1;
            }));
        });
    });
};

test_cases.check_if_simple_render_renders_correctly = function () {
    var spreadsheet = new SpreadSheet();
    return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (valid_template) {
        return spreadsheet.load(new JSZip(valid_template));
    }).then(function (spreadsheet) {
        return spreadsheet._simple_render({ AccountName__c: 'hoge account', AccountAddress__c: 'hoge street' });
    }).then(function (excel_data) {
        var test_spreadsheet = new SpreadSheet();
        return test_spreadsheet.load(new JSZip(excel_data));
    }).then(function (test_spreadsheet) {
        assert(test_spreadsheet.variables.length === 0);
        assert(test_spreadsheet.excel.file('xl/sharedStrings.xml').asText().indexOf('hoge account') !== -1);
        assert(test_spreadsheet.excel.file('xl/sharedStrings.xml').asText().indexOf('hoge street') !== -1);
    });
};

/**
 fs.readFileAsync('./template/Template.xlsx')
 .then((excel_template)=>{
    return Promise.props({
        rendering_data1: readYamlAsync('./data/data1.yml'),     //Load single data
        rendering_data2: readYamlAsync('./data/data2.yml'),     //Load array data
        merge: new ExcelMerge().load(new JSZip(excel_template)) //Initialize ExcelMerge object
    });
})
 */

module.exports = test_cases;