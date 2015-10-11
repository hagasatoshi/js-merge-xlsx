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

module.exports = {
    check_load_with_no_parameter_should_return_error: function check_load_with_no_parameter_should_return_error() {
        var spreadsheet = new SpreadSheet();
        return spreadsheet.load().then(function () {
            throw new Error('test_load_with_no_parameter_should_return_error failed ');
        })['catch'](function (err) {
            assert.equal(err, 'First parameter must be JSZip instance including MS-Excel data');
        });
    },

    check_load_should_return_this_instance: function check_load_should_return_this_instance() {
        var spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (valid_template) {
            return spreadsheet.load(new JSZip(valid_template));
        }).then(function (spreadsheet) {
            assert(spreadsheet instanceof SpreadSheet, 'SpreadSheet#load() should return this instance');
        });
    },

    check_load_each_member_from_valid_template: function check_load_each_member_from_valid_template() {
        var spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (valid_template) {
            return spreadsheet.load(new JSZip(valid_template));
        }).then(function (spreadsheet) {

            //excel
            assert(spreadsheet.excel instanceof JSZip, 'SpreadSheet#excel is not assigned correctly');
            //check if each variables is parsed or not.
            var variables = ['AccountName__c', 'StartDateFormat__c', 'EndDateFormat__c', 'JobDescription__c', 'StartTime__c', 'EndTime__c', 'hasOverTime__c', 'HoliDayType__c', 'Salary__c', 'DueDate__c', 'SalaryDate__c', 'AccountName__c', 'AccountAddress__c'];
            var chk_common_strings_with_variable = _.map(spreadsheet.common_strings_with_variable, function (e) {
                return _(e.t).string_value();
            });
            _.each(variables, function (e) {
                //variables
                assert(_.contains(spreadsheet.variables, e), 'SpreadSheet#load() doesn\'t set up ' + e + ' as variable correctly');
                assert(_.find(chk_common_strings_with_variable, function (v) {
                    return v.indexOf('{{' + e + '}}') !== -1;
                }), 'SpreadSheet#load() doesn\'t set up ' + e + ' as variable correctly');
            });
        });
    },

    simple_render_with_no_parameter_should_return_error: function simple_render_with_no_parameter_should_return_error() {
        var spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (valid_template) {
            return spreadsheet.load(new JSZip(valid_template));
        }).then(function (spreadsheet) {
            return spreadsheet.simple_render();
        }).then(function () {
            throw new Error('simple_render_with_no_parameter_should_return_error failed ');
        })['catch'](function (err) {
            assert.equal(err, 'simple_render() must has parameter');
        });
    },

    check_if_simple_render_renders_correctly: function check_if_simple_render_renders_correctly() {
        var spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (valid_template) {
            return spreadsheet.load(new JSZip(valid_template));
        }).then(function (spreadsheet) {
            return spreadsheet.simple_render({ AccountName__c: 'hoge account', AccountAddress__c: 'hoge street' });
        }).then(function (excel_data) {
            var test_spreadsheet = new SpreadSheet();
            return test_spreadsheet.load(new JSZip(excel_data));
        }).then(function (test_spreadsheet) {
            assert(test_spreadsheet.variables.length === 0, "SpreadSheet#simple_render() doesn't work correctly");
            assert(test_spreadsheet.has_as_shared_string('hoge account'), "'hoge account' is not rendered by SpreadSheet#simple_render()");
            assert(test_spreadsheet.has_as_shared_string('hoge street'), "'hoge street' is not rendered by SpreadSheet#simple_render()");
        });
    },

    bulk_render_multi_file_no_parameter_should_return_error: function bulk_render_multi_file_no_parameter_should_return_error() {
        var spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (valid_template) {
            return spreadsheet.load(new JSZip(valid_template));
        }).then(function (spreadsheet) {
            return spreadsheet.bulk_render_multi_file();
        }).then(function () {
            throw new Error('bulk_render_multi_file_no_parameter_should_return_error failed ');
        })['catch'](function (err) {
            assert.equal(err, 'bulk_render_multi_file() has only array object');
        });
    },

    bulk_render_multi_file_must_have_array_as_parameter: function bulk_render_multi_file_must_have_array_as_parameter() {
        var spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (valid_template) {
            return spreadsheet.load(new JSZip(valid_template));
        }).then(function (spreadsheet) {
            return spreadsheet.bulk_render_multi_file({ name: 'hogehoge' });
        }).then(function () {
            throw new Error('bulk_render_multi_file_must_have_array_as_parameter failed ');
        })['catch'](function (err) {
            assert.equal(err, 'bulk_render_multi_file() has only array object');
        });
    },

    bulk_render_multi_file_must_have_name_and_data: function bulk_render_multi_file_must_have_name_and_data() {
        var spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (valid_template) {
            return spreadsheet.load(new JSZip(valid_template));
        }).then(function (spreadsheet) {
            return spreadsheet.bulk_render_multi_file([{ name: 'hogehoge' }]);
        }).then(function () {
            throw new Error('bulk_render_multi_file_must_have_name_and_data failed ');
        })['catch'](function (err) {
            assert.equal(err, 'bulk_render_multi_file() is called with invalid parameter');
        });
    },

    check_if_bulk_render_multi_file_renders_correctly: function check_if_bulk_render_multi_file_renders_correctly() {
        var spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (valid_template) {
            return spreadsheet.load(new JSZip(valid_template));
        }).then(function (spreadsheet) {
            return spreadsheet.bulk_render_multi_file([{ name: 'file1.xlsx', data: { AccountName__c: 'hoge account1', AccountAddress__c: 'hoge street1' } }, { name: 'file2.xlsx', data: { AccountName__c: 'hoge account2', AccountAddress__c: 'hoge street2' } }, { name: 'file3.xlsx', data: { AccountName__c: 'hoge account3', AccountAddress__c: 'hoge street3' } }]);
        }).then(function (zip_data) {
            var zip = new JSZip(zip_data);
            var excel1 = zip.file('file1.xlsx').asArrayBuffer();
            var excel2 = zip.file('file2.xlsx').asArrayBuffer();
            var excel3 = zip.file('file3.xlsx').asArrayBuffer();
            var spreadsheet_excel1 = new SpreadSheet();
            var spreadsheet_excel2 = new SpreadSheet();
            var spreadsheet_excel3 = new SpreadSheet();
            return Promise.props({
                spreadsheet_excel1: spreadsheet_excel1.load(new JSZip(excel1)),
                spreadsheet_excel2: spreadsheet_excel2.load(new JSZip(excel2)),
                spreadsheet_excel3: spreadsheet_excel3.load(new JSZip(excel3))
            }).then(function (result) {
                var spreadsheet_excel1 = result.spreadsheet_excel1;
                var spreadsheet_excel2 = result.spreadsheet_excel2;
                var spreadsheet_excel3 = result.spreadsheet_excel3;
                assert(spreadsheet_excel1.has_as_shared_string('hoge account1'), "'hoge account1' is missing in excel file");
                assert(spreadsheet_excel1.has_as_shared_string('hoge street1'), "'hoge street1' is missing in excel file");
                assert(spreadsheet_excel2.has_as_shared_string('hoge account2'), "'hoge account2' is missing in excel file");
                assert(spreadsheet_excel2.has_as_shared_string('hoge street2'), "'hoge street2' is missing in excel file");
                assert(spreadsheet_excel3.has_as_shared_string('hoge account3'), "'hoge account3' is missing in excel file");
                assert(spreadsheet_excel3.has_as_shared_string('hoge street3'), "'hoge street3' is missing in excel file");
            });
        });
    },

    add_sheet_binding_data_with_no_parameter_should_return_error: function add_sheet_binding_data_with_no_parameter_should_return_error() {
        var spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (valid_template) {
            return spreadsheet.load(new JSZip(valid_template));
        }).then(function (spreadsheet) {
            return spreadsheet.add_sheet_binding_data();
        }).then(function () {
            throw new Error('add_sheet_binding_data_with_no_parameter_should_return_error failed ');
        })['catch'](function (err) {
            assert.equal(err, 'add_sheet_binding_data() needs to have 2 paramter.');
        });
    },

    add_sheet_binding_data_with_1_parameter_should_return_error: function add_sheet_binding_data_with_1_parameter_should_return_error() {
        var spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (valid_template) {
            return spreadsheet.load(new JSZip(valid_template));
        }).then(function (spreadsheet) {
            return spreadsheet.add_sheet_binding_data('hoge');
        }).then(function () {
            throw new Error('add_sheet_binding_data_with_no_parameter_should_return_error failed ');
        })['catch'](function (err) {
            assert.equal(err, 'add_sheet_binding_data() needs to have 2 paramter.');
        });
    },

    activate_sheet_with_no_parameter_should_return_error: function activate_sheet_with_no_parameter_should_return_error() {
        var spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (valid_template) {
            return spreadsheet.load(new JSZip(valid_template));
        }).then(function (spreadsheet) {
            return spreadsheet.activate_sheet();
        }).then(function () {
            throw new Error('activate_sheet_with_no_parameter_should_return_error failed ');
        })['catch'](function (err) {
            assert.equal(err, 'activate_sheet() needs to have 1 paramter.');
        });
    },

    activate_sheet_with_invalid_sheetname_should_return_error: function activate_sheet_with_invalid_sheetname_should_return_error() {
        var spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (valid_template) {
            return spreadsheet.load(new JSZip(valid_template));
        }).then(function (spreadsheet) {
            return spreadsheet.activate_sheet('hoge');
        }).then(function () {
            throw new Error('activate_sheet_with_no_parameter_should_return_error failed ');
        })['catch'](function (err) {
            assert.equal(err, "Invalid sheet name 'hoge'.");
        });
    },

    delete_sheet_with_no_parameter_should_return_error: function delete_sheet_with_no_parameter_should_return_error() {
        var spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (valid_template) {
            return spreadsheet.load(new JSZip(valid_template));
        }).then(function (spreadsheet) {
            return spreadsheet.delete_sheet();
        }).then(function () {
            throw new Error('delete_sheet_with_no_parameter_should_return_error failed ');
        })['catch'](function (err) {
            assert.equal(err, 'delete_sheet() needs to have 1 paramter.');
        });
    },

    delete_sheet_with_invalid_sheetname_should_return_error: function delete_sheet_with_invalid_sheetname_should_return_error() {
        var spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx').then(function (valid_template) {
            return spreadsheet.load(new JSZip(valid_template));
        }).then(function (spreadsheet) {
            return spreadsheet.delete_sheet('hoge');
        }).then(function () {
            throw new Error('delete_sheet_with_invalid_sheetname_should_return_error failed ');
        })['catch'](function (err) {
            assert.equal(err, "Invalid sheet name 'hoge'.");
        });
    }
};