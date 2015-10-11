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
    it('validation / load() / no parameter / error', test_spreadsheet.check_load_with_no_parameter_should_return_error);
    //it('validation / load() / invalid sheetname / error', test_spreadsheet.check_load_with_invalid_sheetname_should_return_error);
    it('validation / simple_render() / no parameter / error', test_spreadsheet.simple_render_with_no_parameter_should_return_error);
    it('validation / bulk_render_multi_file() / no parameter / error', test_spreadsheet.bulk_render_multi_file_no_parameter_should_return_error);
    it('validation / bulk_render_multi_file() / object / error', test_spreadsheet.bulk_render_multi_file_must_have_array_as_parameter);
    it('validation / bulk_render_multi_file() / object / error', test_spreadsheet.bulk_render_multi_file_must_have_name_and_data);
    it('validation / add_sheet_binding_data() / no parameter / error', test_spreadsheet.add_sheet_binding_data_with_no_parameter_should_return_error);
    it('validation / add_sheet_binding_data() / 1 parameter / error', test_spreadsheet.add_sheet_binding_data_with_1_parameter_should_return_error);
    it('validation / activate_sheet() / no parameter / error', test_spreadsheet.activate_sheet_with_no_parameter_should_return_error);
    it('validation / activate_sheet() / invalid sheetname / error', test_spreadsheet.activate_sheet_with_invalid_sheetname_should_return_error);
    it('validation / delete_sheet() / no parameter / error', test_spreadsheet.delete_sheet_with_no_parameter_should_return_error);
    it('validation / delete_sheet() / invalid sheetname / error', test_spreadsheet.delete_sheet_with_invalid_sheetname_should_return_error);

    /* test for logic */
    it('logic / load() / load each member from valid template', test_spreadsheet.check_load_each_member_from_valid_template);
    it('logic / load() / should return this instance', test_spreadsheet.check_load_should_return_this_instance);
    it('logic / simple_render() / renders correctly', test_spreadsheet.check_if_simple_render_renders_correctly);
    it('logic / bulk_render_multi_file() / renders correctly', test_spreadsheet.check_if_bulk_render_multi_file_renders_correctly);

});


/** output test */
describe('output test : ',  ()=>{
    let util = new Utilitly();

    /* output test */
    it('single / normaldata / Template.xlsx', ()=>util.output('Template.xlsx','single_normal_data.yml',EXCEL_OUTPUT_TYPE.SINGLE,'single_normaldata_temlate.xlsx'));
    it('bulk / normaldata / Template.xlsx', ()=>util.output('Template.xlsx','bulk_normal_data.yml',EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_FILE,'bulk_normaldata_temlate.zip'));
    it('bulk / normaldata / Template.xlsx', ()=>util.output('Template.xlsx','bulk_normal_data.yml',EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_SHEET,'bulk_normaldata_temlate.xlsx'));
});

