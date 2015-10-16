/**
 * * test.js
 * * Test script for js-merge-xlsx
 * * @author Satoshi Haga
 * * @date 2015/09/30
 **/
//require test cases
var assert = require('assert');
var test_spreadsheet = require('./test_cases/test_spreadsheet');
var Utilitly = require('./test_cases/test_output_excel_file');

var EXCEL_OUTPUT_TYPE = {
    SINGLE : 0,
    BULK_MULTIPLE_FILE : 1,
    BULK_MULTIPLE_SHEET : 2
};

/** test for spreadsheet.js */
describe('Test for spreadsheet.js : ',  ()=>{

    /* test for validation */
    it('validation / load() / no parameter / error', test_spreadsheet.checkLoadWithNoParameterShouldReturnError);
    it('validation / simpleRender() / no parameter / error', test_spreadsheet.simpleMergeWithNoParameterShouldReturnError);
    it('validation / bulkMergeMultiFile() / no parameter / error', test_spreadsheet.bulkMergeMultiFileNoParameterShouldReturnError);
    it('validation / bulkMergeMultiFile() / object / error', test_spreadsheet.bulkMergeMultiFileMustHaveArrayAsParameter);
    it('validation / bulkMergeMultiFile() / object / error', test_spreadsheet.bulkMergeMultiFileMustHaveNameAndData);
    it('validation / addSheetBindingData() / no parameter / error', test_spreadsheet.addSheetBindingDataWithNoParameterShouldReturnError);
    it('validation / addSheetBindingData() / 1 parameter / error', test_spreadsheet.addSheetBindingDataWith1ParameterShouldReturnError);
    it('validation / activateSheet() / no parameter / error', test_spreadsheet.activateSheetWithNoParameterShouldReturnError);
    it('validation / activateSheet() / invalid sheetname / error', test_spreadsheet.activateSheetWithInvalidSheetnameShouldReturnError);
    it('validation / deleteSheet() / no parameter / error', test_spreadsheet.deleteSheetWithNoParameterShouldReturnError);
    it('validation / deleteSheet() / invalid sheetname / error', test_spreadsheet.deleteSheetWithInvalidSheetnameShouldReturnError);

    /* test for logic */
    it('logic / load() / load each member from valid template', test_spreadsheet.checkLoadEachMemberFromValidTemplate);
    it('logic / load() / should return this instance', test_spreadsheet.checkLoadShouldReturnThisInstance);
    it('logic / simpleRender() / renders correctly', test_spreadsheet.checkIfSimpleMergeRendersCorrectly);
    it('logic / bulkMergeMultiFile() / renders correctly', test_spreadsheet.checkIfBulkMergeMultiFileRendersCorrectly);

});


/** output test */
describe('output test : ',  ()=>{
    let util = new Utilitly();

    /* 01.Normal Case */
    it('single / normaldata / Template.xlsx',
        ()=>util.output('Template.xlsx','single_normal_data.yml',EXCEL_OUTPUT_TYPE.SINGLE,'01_normal_case/single_normaldata_temlate.xlsx'));
    it('bulk / normaldata / Template.xlsx',
        ()=>util.output('Template.xlsx','bulk_normal_data.yml',EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_FILE,'01_normal_case/bulk_normaldata_temlate.zip'));
    it('bulk / normaldata / Template.xlsx',
        ()=>util.output('Template.xlsx','bulk_normal_data.yml',EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_SHEET,'01_normal_case/bulk_normaldata_temlate.xlsx'));

    /* 02.No Image */
    it('single / normaldata / TemplateNoImage.xlsx',
        ()=>util.output('TemplateNoImage.xlsx','single_normal_data.yml',EXCEL_OUTPUT_TYPE.SINGLE,'02_no_image/single_normaldata_noimage.xlsx'));
    it('bulk / normaldata / TemplateNoImage.xlsx',
        ()=>util.output('TemplateNoImage.xlsx','bulk_normal_data.yml',EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_FILE,'02_no_image/bulk_normaldata_noimage.zip'));
    it('bulk / normaldata / TemplateNoImage.xlsx',
        ()=>util.output('TemplateNoImage.xlsx','bulk_normal_data.yml',EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_SHEET,'02_no_image/bulk_normaldata_noimage.xlsx'));

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

    /* 04.String with no variables */
    it('single / normaldata / TemplateStringWithNoVariables.xlsx', function () {
        return util.output('TemplateStringWithNoVariables.xlsx', 'single_normal_data.yml', EXCEL_OUTPUT_TYPE.SINGLE, '04_variables/single_normaldata_novariables.xlsx');
    });
    it('bulk / normaldata / TemplateStringWithNoVariables.xlsx', function () {
        return util.output('TemplateStringWithNoVariables.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_FILE, '04_variables/bulk_normaldata_novariables.zip');
    });
    it('bulk / normaldata / TemplateStringWithNoVariables.xlsx', function () {
        return util.output('TemplateStringWithNoVariables.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_SHEET, '04_variables/bulk_normaldata_novariables.xlsx');
    });

    /* 05.multiple images */
    it('single / normaldata / TemplateMultiImages.xlsx', function () {
        return util.output('TemplateMultiImages.xlsx', 'single_normal_data.yml', EXCEL_OUTPUT_TYPE.SINGLE, '05_multiple_images/single_normaldata_multiple_images.xlsx');
    });
    it('bulk / normaldata / TemplateMultiImages.xlsx', function () {
        return util.output('TemplateMultiImages.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_FILE, '05_multiple_images/bulk_normaldata_multiple_images.zip');
    });
    it('bulk / normaldata / TemplateMultiImages.xlsx', function () {
        return util.output('TemplateMultiImages.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_SHEET, '05_multiple_images/bulk_normaldata_multiple_images.xlsx');
    });

    /* 06.With objects */
    it('single / normaldata / TemplateWithObject.xlsx', function () {
        return util.output('TemplateWithObject.xlsx', 'single_normal_data.yml', EXCEL_OUTPUT_TYPE.SINGLE, '06_object/single_normaldata_with_object.xlsx');
    });
    it('bulk / normaldata / TemplateWithObject.xlsx', function () {
        return util.output('TemplateWithObject.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_FILE, '06_object/bulk_normaldata_with_object.zip');
    });
    it('bulk / normaldata / TemplateWithObject.xlsx', function () {
        return util.output('TemplateWithObject.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_SHEET, '06_object/bulk_normaldata_with_object.xlsx');
    });

});
