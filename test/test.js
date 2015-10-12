/**
 * * test.js
 * * Test script for js-merge-xlsx
 * * @author Satoshi Haga
 * * @date 2015/09/30
 **/
//require test cases
'use strict';

var assert = require('assert');
var test_spreadsheet = require('./test_cases/test_spreadsheet');
var Utilitly = require('./test_cases/test_output_excel_file');

var EXCEL_OUTPUT_TYPE = {
    SINGLE: 0,
    BULK_MULTIPLE_FILE: 1,
    BULK_MULTIPLE_SHEET: 2
};

/** test for spreadsheet.js */
describe('Test for spreadsheet.js : ', function () {
    /* test for validation */
    it('validation / load() / no parameter / error', test_spreadsheet.checkLoadWithNoParameterShouldReturnError);
    it('validation / simpleRender() / no parameter / error', test_spreadsheet.simpleRenderWithNoParameterShouldReturnError);
    it('validation / bulkRender_multi_file() / no parameter / error', test_spreadsheet.bulkRenderMultiFileNoParameterShouldReturnError);
    it('validation / bulkRender_multi_file() / object / error', test_spreadsheet.bulkRenderMultiFileMustHaveArrayAsParameter);
    it('validation / bulkRender_multi_file() / object / error', test_spreadsheet.bulkRenderMultiFileMustHaveNameAndData);
    it('validation / addSheet_binding_data() / no parameter / error', test_spreadsheet.addSheetBindingDataWithNoParameterShouldReturnError);
    it('validation / addSheet_binding_data() / 1 parameter / error', test_spreadsheet.addSheetBindingDataWith1ParameterShouldReturnError);
    it('validation / activateSheet() / no parameter / error', test_spreadsheet.activateSheetWithNoParameterShouldReturnError);
    it('validation / activateSheet() / invalid sheetname / error', test_spreadsheet.activateSheetWithInvalidSheetnameShouldReturnError);
    it('validation / deleteSheet() / no parameter / error', test_spreadsheet.deleteSheetWithNoParameterShouldReturnError);
    it('validation / deleteSheet() / invalid sheetname / error', test_spreadsheet.deleteSheetWithInvalidSheetnameShouldReturnError);
    /* test for logic */
    it('logic / load() / load each member from valid template', test_spreadsheet.checkLoadEachMemberFromValidTemplate);
    it('logic / load() / should return this instance', test_spreadsheet.checkLoadShouldReturnThisInstance);
    it('logic / simpleRender() / renders correctly', test_spreadsheet.checkIfSimpleRenderRendersCorrectly);
    it('logic / bulkRenderMultiFile() / renders correctly', test_spreadsheet.checkIfBulkRenderMultiFileRendersCorrectly);
});

/** output test */
describe('output test : ', function () {
    var util = new Utilitly();

    /* output test */
    it('single / normaldata / Template.xlsx', function () {
        return util.output('Template.xlsx', 'single_normal_data.yml', EXCEL_OUTPUT_TYPE.SINGLE, 'single_normaldata_temlate.xlsx');
    });
    it('bulk / normaldata / Template.xlsx', function () {
        return util.output('Template.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_FILE, 'bulk_normaldata_temlate.zip');
    });
    it('bulk / normaldata / Template.xlsx', function () {
        return util.output('Template.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_SHEET, 'bulk_normaldata_temlate.xlsx');
    });
});