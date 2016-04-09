/**
 * test.js
 * Test script for js-merge-xlsx
 * @author Satoshi Haga
 * @date 2015/09/30
 */
'use strict';

var assert = require('assert');
var test_spreadsheet = require('./test_cases/test_spreadsheet');
var test_excelmerge = require('./test_cases/test_excelmerge');
var Utilitly = require('./test_cases/test_output_excel_file');

var EXCEL_OUTPUT_TYPE = {
    SINGLE: 0,
    BULK_MULTIPLE_FILE: 1,
    BULK_MULTIPLE_SHEET: 2
};

/** SpreadSheet */
describe('Test for spreadsheet.js : ', function () {
    it('logic / load() / load each member from valid template', test_spreadsheet.checkLoadEachMemberFromValidTemplate);
    it('logic / load() / should return this instance', test_spreadsheet.checkLoadShouldReturnThisInstance);
    it('logic / simpleMerge() / renders correctly', test_spreadsheet.checkIfSimpleMergeRendersCorrectly);
    it('logic / bulkMergeMultiFile() / renders correctly', test_spreadsheet.checkIfBulkMergeMultiFileRendersCorrectly);
    it('logic / addSheetBindingData() / works correctly', test_spreadsheet.checkIfAddSheetBindingDataCorrectly);
    it('logic / deleteTemplateSheet() / works correctly', test_spreadsheet.checkIfDeleteTemplateSheetWorksCorrectly);
    it('logic / templateVariablesWorkCorrectly() / works correctly', test_spreadsheet.checkTemplateVariablesWorkCorrectly);
});

/** ExcelMerge */
describe('Test for ExcelMerge.js : ', function () {

    //Validation
    it('validation / load() / no parameter / error', test_excelmerge.checkLoadWithNoParameterShouldReturnError);
    it('logic / merge() with no parameter / should return error', test_excelmerge.checkIfMergeWithNoParameterRendersCorrectly);
    it('logic / bulkMergeMultiFile() with no parameter / should return error', test_excelmerge.checkIfBulkMergeMultiFileWithNoParameterShouldReturnError);
    it('logic / bulkMergeMultiSheet() with no parameter / should return error', test_excelmerge.checkIfBulkMergeMultiSheetWithNoParameterShouldReturnError);

    //Core
    it('logic / load() / load each member from valid template', test_excelmerge.checkLoadEachMemberFromValidTemplate);
    it('logic / load() / should return this instance', test_excelmerge.checkLoadShouldReturnThisInstance);
    it('logic / merge() / renders correctly', test_excelmerge.checkIfMergeRendersCorrectly);
    it('logic / mergeByType(SINGLE_DATA) / renders correctly', test_excelmerge.checkIfMergeByTypeRendersCorrectly1);
    it('logic / bulkMergeMultiFile() / renders correctly', test_excelmerge.checkIfBulkMergeMultiFileRendersCorrectly);
    it('logic / mergeByType(MULTI_FILE) / renders correctly', test_excelmerge.checkIfMergeByTypeRendersCorrectly2);
    it('logic / bulkMergeMultiSheet() / renders correctly', test_excelmerge.checkIfBulkMergeMultiSheetRendersCorrectly);
    it('logic / mergeByType(MULTI_SHEET) / renders correctly', test_excelmerge.checkIfMergeByTypeRendersCorrectly3);
    it('logic / variables() / parse correctly', test_excelmerge.checkVariablesWorkCorrectly);
    it('logic / mergeByType() / throw error with invalid parameter', test_excelmerge.checkIfMergeByTypeThrowErrorWithInvalidType);
});

/** Output test */
describe('output test : ', function () {
    var util = new Utilitly();

    //01.Normal Case
    it('single / normaldata / Template.xlsx', function () {
        return util.output('Template.xlsx', 'single_normal_data.yml', EXCEL_OUTPUT_TYPE.SINGLE, '01_normal_case/single_normaldata_temlate.xlsx');
    });
    it('bulk / normaldata / Template.xlsx', function () {
        return util.output('Template.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_FILE, '01_normal_case/bulk_normaldata_temlate.zip');
    });
    it('bulk / normaldata / Template.xlsx', function () {
        return util.output('Template.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_SHEET, '01_normal_case/bulk_normaldata_temlate.xlsx');
    });

    //02.No Image
    it('single / normaldata / TemplateNoImage.xlsx', function () {
        return util.output('TemplateNoImage.xlsx', 'single_normal_data.yml', EXCEL_OUTPUT_TYPE.SINGLE, '02_no_image/single_normaldata_noimage.xlsx');
    });
    it('bulk / normaldata / TemplateNoImage.xlsx', function () {
        return util.output('TemplateNoImage.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_FILE, '02_no_image/bulk_normaldata_noimage.zip');
    });
    it('bulk / normaldata / TemplateNoImage.xlsx', function () {
        return util.output('TemplateNoImage.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_SHEET, '02_no_image/bulk_normaldata_noimage.xlsx');
    });

    //03.No String
    it('single / normaldata / TemplateNoStrings.xlsx', function () {
        return util.output('TemplateNoStrings.xlsx', 'single_normal_data.yml', EXCEL_OUTPUT_TYPE.SINGLE, '03_no_strings/single_normaldata_nostrings.xlsx');
    });
    it('bulk / normaldata / TemplateNoStrings.xlsx', function () {
        return util.output('TemplateNoStrings.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_FILE, '03_no_strings/bulk_normaldata_nostrings.zip');
    });
    it('bulk / normaldata / TemplateNoStrings.xlsx', function () {
        return util.output('TemplateNoStrings.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_SHEET, '03_no_strings/bulk_normaldata_nostrings.xlsx');
    });

    //04.String with no variables
    it('single / normaldata / TemplateStringWithNoVariables.xlsx', function () {
        return util.output('TemplateStringWithNoVariables.xlsx', 'single_normal_data.yml', EXCEL_OUTPUT_TYPE.SINGLE, '04_variables/single_normaldata_novariables.xlsx');
    });
    it('bulk / normaldata / TemplateStringWithNoVariables.xlsx', function () {
        return util.output('TemplateStringWithNoVariables.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_FILE, '04_variables/bulk_normaldata_novariables.zip');
    });
    it('bulk / normaldata / TemplateStringWithNoVariables.xlsx', function () {
        return util.output('TemplateStringWithNoVariables.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_SHEET, '04_variables/bulk_normaldata_novariables.xlsx');
    });

    //05.multiple images
    it('single / normaldata / TemplateMultiImages.xlsx', function () {
        return util.output('TemplateMultiImages.xlsx', 'single_normal_data.yml', EXCEL_OUTPUT_TYPE.SINGLE, '05_multiple_images/single_normaldata_multiple_images.xlsx');
    });
    it('bulk / normaldata / TemplateMultiImages.xlsx', function () {
        return util.output('TemplateMultiImages.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_FILE, '05_multiple_images/bulk_normaldata_multiple_images.zip');
    });
    it('bulk / normaldata / TemplateMultiImages.xlsx', function () {
        return util.output('TemplateMultiImages.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_SHEET, '05_multiple_images/bulk_normaldata_multiple_images.xlsx');
    });

    //06.With objects
    it('single / normaldata / TemplateWithObject.xlsx', function () {
        return util.output('TemplateWithObject.xlsx', 'single_normal_data.yml', EXCEL_OUTPUT_TYPE.SINGLE, '06_object/single_normaldata_with_object.xlsx');
    });
    it('bulk / normaldata / TemplateWithObject.xlsx', function () {
        return util.output('TemplateWithObject.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_FILE, '06_object/bulk_normaldata_with_object.zip');
    });
    it('bulk / normaldata / TemplateWithObject.xlsx', function () {
        return util.output('TemplateWithObject.xlsx', 'bulk_normal_data.yml', EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_SHEET, '06_object/bulk_normaldata_with_object.xlsx');
    });

    //07.Character Test
    it('single / characterdata / Template.xlsx', function () {
        return util.output_character_test_single_record('Template.xlsx', '07_character/single_character.xlsx');
    });
    it('bulk / characterdata / Template.xlsx', function () {
        return util.output_character_test_bulk_record_as_multifile('Template.xlsx', '07_character/bulk_character.zip');
    });
    it('bulk / characterdata / Template.xlsx', function () {
        return util.output_character_test_bulk_record_as_multisheet('Template.xlsx', '07_character/bulk_character.xlsx');
    });
});