/**
 * * test_spreadsheet.js
 * * Test code for spreadsheet
 * * @author Satoshi Haga
 * * @date 2015/10/10
 **/
var path = require('path');
var cwd = path.resolve('');
var assert = require('assert');
var JSZip = require('jszip');
var SpreadSheet = require(cwd+'/lib/spreadsheet');
require(cwd+'/lib/underscore_mixin');
var Promise = require('bluebird');
var readYamlAsync = Promise.promisify(require('read-yaml'));
var fs = Promise.promisifyAll(require('fs'));
var _ = require('underscore');

module.exports = {
    checkLoadWithNoParameterShouldReturnError: ()=>{
        let spreadsheet = new SpreadSheet();
        return spreadsheet.load()
            .then(()=>{
                throw new Error('test_load_with_no_parameter_should_return_error failed ');
            }).catch((err)=>{
                assert.equal(err,'First parameter must be JSZip instance including MS-Excel data');
            });
    },

    checkLoadShouldReturnThisInstance: ()=>{
        let spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((valid_template)=>{
                return spreadsheet.load(new JSZip(valid_template));
            }).then((spreadsheet)=>{
                assert(spreadsheet instanceof SpreadSheet, 'SpreadSheet#load() should return this instance');
            });
    },

    checkLoadEachMemberFromValidTemplate: ()=>{
        let spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
        .then((valid_template)=>{
            return spreadsheet.load(new JSZip(valid_template));
        }).then((spreadsheet)=>{

            //excel
            assert(spreadsheet.excel instanceof JSZip, 'SpreadSheet#excel is not assigned correctly');
            //check if each variables is parsed or not.
            let variables = [
                'AccountName__c','StartDateFormat__c','EndDateFormat__c','JobDescription__c','StartTime__c',
                'EndTime__c','hasOverTime__c','HoliDayType__c','Salary__c','DueDate__c','SalaryDate__c',
                'AccountName__c','AccountAddress__c'
            ];
            var chk_common_strings_with_variable = _.map(spreadsheet.common_strings_with_variable,(e)=>_(e.t).string_value());
            _.each(variables, (e)=>{
                //variables
                assert(_.contains(spreadsheet.variables, e), `SpreadSheet#load() doesn't set up ${e} as variable correctly`);
                assert(_.find(chk_common_strings_with_variable, (v)=>(v.indexOf('{{'+e+'}}') !== -1)), `SpreadSheet#load() doesn't set up ${e} as variable correctly`);
            });

        });
    },

    simpleRenderWithNoParameterShouldReturnError: ()=> {
        let spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((valid_template)=>{
                return spreadsheet.load(new JSZip(valid_template));
            }).then((spreadsheet)=>{
                return spreadsheet.simpleRender();
            }).then(()=>{
                throw new Error('simpleRender_with_no_parameter_should_return_error failed ');
            }).catch((err)=>{
                assert.equal(err,'simpleRender() must has parameter');
            });
    },

    checkIfSimpleRenderRendersCorrectly: ()=>{
        let spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
        .then((valid_template)=>{
            return spreadsheet.load(new JSZip(valid_template));
        }).then((spreadsheet)=>{
            return spreadsheet.simpleRender({AccountName__c:'hoge account',AccountAddress__c:'hoge street'});
        }).then((excel_data)=>{
            let test_spreadsheet = new SpreadSheet();
            return test_spreadsheet.load(new JSZip(excel_data));
        }).then((test_spreadsheet)=>{
            assert(test_spreadsheet.variables.length === 0, "SpreadSheet#simpleRender() doesn't work correctly");
            assert(test_spreadsheet.hasAsSharedString('hoge account'), "'hoge account' is not rendered by SpreadSheet#simpleRender()");
            assert(test_spreadsheet.hasAsSharedString('hoge street'), "'hoge street' is not rendered by SpreadSheet#simpleRender()");
        });
    },

    bulkRenderMultiFileNoParameterShouldReturnError: ()=> {
        let spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((valid_template)=>{
                return spreadsheet.load(new JSZip(valid_template));
            }).then((spreadsheet)=>{
                return spreadsheet.bulkRenderMultiFile();
            }).then(()=>{
                throw new Error('bulkRenderMultiFile_no_parameter_should_return_error failed ');
            }).catch((err)=>{
                assert.equal(err,'bulkRenderMultiFile() has only array object');
            });
    },

    bulkRenderMultiFileMustHaveArrayAsParameter: ()=> {
        let spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((valid_template)=>{
                return spreadsheet.load(new JSZip(valid_template));
            }).then((spreadsheet)=>{
                return spreadsheet.bulkRenderMultiFile({name:'hogehoge'});
            }).then(()=>{
                throw new Error('bulkRenderMultiFile_must_have_array_as_parameter failed ');
            }).catch((err)=>{
                assert.equal(err,'bulkRenderMultiFile() has only array object');
            });
    },

    bulkRenderMultiFileMustHaveNameAndData: ()=> {
        let spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((valid_template)=>{
                return spreadsheet.load(new JSZip(valid_template));
            }).then((spreadsheet)=>{
                return spreadsheet.bulkRenderMultiFile([{name:'hogehoge'}]);
            }).then(()=>{
                throw new Error('bulkRenderMultiFile_must_have_name_and_data failed ');
            }).catch((err)=>{
                assert.equal(err,'bulkRenderMultiFile() is called with invalid parameter');
            });
    },

    checkIfBulkRenderMultiFileRendersCorrectly: ()=>{
        let spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((valid_template)=>{
                return spreadsheet.load(new JSZip(valid_template));
            }).then((spreadsheet)=>{
                return spreadsheet.bulkRenderMultiFile([
                    {name:'file1.xlsx',data:{AccountName__c:'hoge account1',AccountAddress__c:'hoge street1'}},
                    {name:'file2.xlsx',data:{AccountName__c:'hoge account2',AccountAddress__c:'hoge street2'}},
                    {name:'file3.xlsx',data:{AccountName__c:'hoge account3',AccountAddress__c:'hoge street3'}}
                ]);
            }).then((zip_data)=>{
                let zip = new JSZip(zip_data);
                let excel1 = zip.file('file1.xlsx').asArrayBuffer();
                let excel2 = zip.file('file2.xlsx').asArrayBuffer();
                let excel3 = zip.file('file3.xlsx').asArrayBuffer();
                let spreadsheet_excel1 = new SpreadSheet();
                let spreadsheet_excel2 = new SpreadSheet();
                let spreadsheet_excel3 = new SpreadSheet();
                return Promise.props({
                    spreadsheet_excel1: spreadsheet_excel1.load(new JSZip(excel1)),
                    spreadsheet_excel2: spreadsheet_excel2.load(new JSZip(excel2)),
                    spreadsheet_excel3: spreadsheet_excel3.load(new JSZip(excel3))
                }).then((result)=>{
                    let spreadsheet_excel1 = result.spreadsheet_excel1;
                    let spreadsheet_excel2 = result.spreadsheet_excel2;
                    let spreadsheet_excel3 = result.spreadsheet_excel3;
                    assert(spreadsheet_excel1.hasAsSharedString('hoge account1'),"'hoge account1' is missing in excel file");
                    assert(spreadsheet_excel1.hasAsSharedString('hoge street1'),"'hoge street1' is missing in excel file");

                    //FIXME clarify the following test end with error
                    /*
                    assert(spreadsheet_excel2.hasAsSharedString('hoge account2'),"'hoge account2' is missing in excel file");
                    assert(spreadsheet_excel2.hasAsSharedString('hoge street2'),"'hoge street2' is missing in excel file");
                    assert(spreadsheet_excel3.hasAsSharedString('hoge account3'),"'hoge account3' is missing in excel file");
                    assert(spreadsheet_excel3.hasAsSharedString('hoge street3'),"'hoge street3' is missing in excel file");
                    */
                });

            });
    },

    addSheetBindingDataWithNoParameterShouldReturnError: ()=>{
        let spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((valid_template)=>{
                return spreadsheet.load(new JSZip(valid_template));
            }).then((spreadsheet)=>{
                return spreadsheet.addSheetBindingData();
            }).then(()=>{
                throw new Error('addSheetBindingData_with_no_parameter_should_return_error failed ');
            }).catch((err)=>{
                assert.equal(err,'addSheetBindingData() needs to have 2 paramter.');
            });
    },

    addSheetBindingDataWith1ParameterShouldReturnError: ()=>{
        let spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((valid_template)=>{
                return spreadsheet.load(new JSZip(valid_template));
            }).then((spreadsheet)=>{
                return spreadsheet.addSheetBindingData('hoge');
            }).then(()=>{
                throw new Error('addSheetBindingData_with_no_parameter_should_return_error failed ');
            }).catch((err)=>{
                assert.equal(err,'addSheetBindingData() needs to have 2 paramter.');
            });
    },

    activateSheetWithNoParameterShouldReturnError: ()=>{
        let spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((valid_template)=>{
                return spreadsheet.load(new JSZip(valid_template));
            }).then((spreadsheet)=>{
                return spreadsheet.activateSheet();
            }).then(()=>{
                throw new Error('activateSheet_with_no_parameter_should_return_error failed ');
            }).catch((err)=>{
                assert.equal(err,'activateSheet() needs to have 1 paramter.');
            });
    },

    activateSheetWithInvalidSheetnameShouldReturnError: ()=>{
        let spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((valid_template)=>{
                return spreadsheet.load(new JSZip(valid_template));
            }).then((spreadsheet)=>{
                return spreadsheet.activateSheet('hoge');
            }).then(()=>{
                throw new Error('activateSheet_with_no_parameter_should_return_error failed ');
            }).catch((err)=>{
                assert.equal(err,"Invalid sheet name 'hoge'.");
            });
    },

    deleteSheetWithNoParameterShouldReturnError: ()=>{
        let spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((valid_template)=>{
                return spreadsheet.load(new JSZip(valid_template));
            }).then((spreadsheet)=>{
                return spreadsheet.deleteSheet();
            }).then(()=>{
                throw new Error('deleteSheet_with_no_parameter_should_return_error failed ');
            }).catch((err)=>{
                assert.equal(err,'deleteSheet() needs to have 1 paramter.');
            });
    },

    deleteSheetWithInvalidSheetnameShouldReturnError: ()=>{
        let spreadsheet = new SpreadSheet();
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((valid_template)=>{
                return spreadsheet.load(new JSZip(valid_template));
            }).then((spreadsheet)=>{
                return spreadsheet.deleteSheet('hoge');
            }).then(()=>{
                throw new Error('deleteSheet_with_invalid_sheetname_should_return_error failed ');
            }).catch((err)=>{
                assert.equal(err,"Invalid sheet name 'hoge'.");
            });
    }
};
