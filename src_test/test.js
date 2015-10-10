/**
 * * test.js
 * * Test script for js-merge-xlsx
 * * @author Satoshi Haga
 * * @date 2015/09/30
 **/

var test_spreadsheet = require('./test_cases/test_spreadsheet');

/** test for spreadsheet */
describe('Test for spreadsheet : ',  ()=>{

    it('load() with no parameter should return error',
        test_spreadsheet.check_load_with_no_parameter_should_return_error);

    it('load each member from valid template',
        test_spreadsheet.check_load_each_member_from_valid_template);

    it('load() should return this instance',
        test_spreadsheet.check_load_should_return_this_instance);

    it('load() with invalid sheetname should return error',
        test_spreadsheet.check_load_with_invalid_sheetname_should_return_error);

    it('check if simple_render() renders correctly',
        test_spreadsheet.check_if_simple_render_renders_correctly);

});
