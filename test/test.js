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
    it('validation / bulkRenderMultiFile() / no parameter / error', test_spreadsheet.bulkRenderMultiFileNoParameterShouldReturnError);
    it('validation / bulkRenderMultiFile() / object / error', test_spreadsheet.bulkRenderMultiFileMustHaveArrayAsParameter);
    it('validation / bulkRenderMultiFile() / object / error', test_spreadsheet.bulkRenderMultiFileMustHaveNameAndData);
    it('validation / addSheetBindingData() / no parameter / error', test_spreadsheet.addSheetBindingDataWithNoParameterShouldReturnError);
    it('validation / addSheetBindingData() / 1 parameter / error', test_spreadsheet.addSheetBindingDataWith1ParameterShouldReturnError);
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

    /* 01.Normal Case */
    it('single / normaldata / Template.xlsx', function () {
        return util.output('Template.xlsx', 'single_normal_data.yml', EXCEL_OUTPUT_TYPE.SINGLE, '01_normal_case/single_normaldata_temlate.xlsx');
    });
    it('bulk / normaldata / Template.xlsx', function () {
        return util.output('Template.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_FILE, '01_normal_case/bulk_normaldata_temlate.zip');
    });
    it('bulk / normaldata / Template.xlsx', function () {
        return util.output('Template.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_SHEET, '01_normal_case/bulk_normaldata_temlate.xlsx');
    });

    /* 02.No Image */
    it('single / normaldata / TemplateNoImage.xlsx', function () {
        return util.output('TemplateNoImage.xlsx', 'single_normal_data.yml', EXCEL_OUTPUT_TYPE.SINGLE, '02_no_image/single_normaldata_noimage.xlsx');
    });
    it('bulk / normaldata / TemplateNoImage.xlsx', function () {
        return util.output('TemplateNoImage.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_FILE, '02_no_image/bulk_normaldata_noimage.zip');
    });
    it('bulk / normaldata / TemplateNoImage.xlsx', function () {
        return util.output('TemplateNoImage.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_SHEET, '02_no_image/bulk_normaldata_noimage.xlsx');
    });

    /* 03.No String */
    it('single / normaldata / TemplateNoStrings.xlsx', function () {
        return util.output('TemplateNoStrings.xlsx', 'single_normal_data.yml', EXCEL_OUTPUT_TYPE.SINGLE, '03_no_strings/single_normaldata_nostrings.xlsx');
    });
    it('bulk / normaldata / TemplateNoStrings.xlsx', function () {
        return util.output('TemplateNoStrings.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_FILE, '03_no_strings/bulk_normaldata_nostrings.zip');
    });
    it('bulk / normaldata / TemplateNoStrings.xlsx', function () {
        return util.output('TemplateNoStrings.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_SHEET, '03_no_strings/bulk_normaldata_nostrings.xlsx');
    });
});