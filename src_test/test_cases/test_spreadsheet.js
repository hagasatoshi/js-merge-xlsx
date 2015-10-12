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
        return new SpreadSheet().load()
            .then(()=>{
                throw new Error('test_load_with_no_parameter_should_return_error failed ');
            }).catch((err)=>{
                assert.equal(err,'First parameter must be JSZip instance including MS-Excel data');
            });
    },

    checkLoadShouldReturnThisInstance: ()=>{
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((validTemplate)=>{
                return new SpreadSheet().load(new JSZip(validTemplate));
            }).then((spreadsheet)=>{
                assert(spreadsheet instanceof SpreadSheet, 'SpreadSheet#load() should return this instance');
            });
    },

    checkLoadEachMemberFromValidTemplate: ()=>{
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((validTemplate)=>{
                return new SpreadSheet().load(new JSZip(validTemplate));
            }).then((spreadsheet)=>{

                //excel
                assert(spreadsheet.excel instanceof JSZip, 'SpreadSheet#excel is not assigned correctly');

                //check if each variables is parsed or not.
                let variables = [
                    'AccountName__c','StartDateFormat__c','EndDateFormat__c','JobDescription__c','StartTime__c',
                    'EndTime__c','hasOverTime__c','HoliDayType__c','Salary__c','DueDate__c','SalaryDate__c',
                    'AccountName__c','AccountAddress__c'
                ];
                var chkCommonStringsWithVariable = _.map(spreadsheet.commonStringsWithVariable,(e)=>_(e.t).stringValue());
                _.each(variables, (e)=>{
                    //variables
                    assert(_.contains(spreadsheet.variables, e), `SpreadSheet#load() doesn't set up ${e} as variable correctly`);
                    assert(_.find(chkCommonStringsWithVariable, (v)=>(v.indexOf(`{{${e}}}`) !== -1)), `SpreadSheet#load() doesn't set up ${e} as variable correctly`);
                });

            });
    },

    simpleRenderWithNoParameterShouldReturnError: ()=> {
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((validTemplate)=>{
                return new SpreadSheet().load(new JSZip(validTemplate));
            }).then((spreadsheet)=>{
                return spreadsheet.simpleRender();
            }).then(()=>{
                throw new Error('simpleRender_with_no_parameter_should_return_error failed ');
            }).catch((err)=>{
                assert.equal(err,'simpleRender() must has parameter');
            });
    },

    checkIfSimpleRenderRendersCorrectly: ()=>{
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((validTemplate)=>{
                return new SpreadSheet().load(new JSZip(validTemplate));
            }).then((spreadsheet)=>{
                return spreadsheet.simpleRender({AccountName__c:'hoge account',AccountAddress__c:'hoge street'});
            }).then((excelData)=>{
                return new SpreadSheet().load(new JSZip(excelData));
            }).then((spreadsheet)=>{
                assert(spreadsheet.variables.length === 0, "SpreadSheet#simpleRender() doesn't work correctly");
                assert(spreadsheet.hasAsSharedString('hoge account'), "'hoge account' is not rendered by SpreadSheet#simpleRender()");
                assert(spreadsheet.hasAsSharedString('hoge street'), "'hoge street' is not rendered by SpreadSheet#simpleRender()");
            });
    },

    bulkRenderMultiFileNoParameterShouldReturnError: ()=> {
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((validTemplate)=>{
                return new SpreadSheet().load(new JSZip(validTemplate));
            }).then((spreadsheet)=>{
                return spreadsheet.bulkRenderMultiFile();
            }).then(()=>{
                throw new Error('bulkRenderMultiFile_no_parameter_should_return_error failed ');
            }).catch((err)=>{
                assert.equal(err,'bulkRenderMultiFile() has only array object');
            });
    },

    bulkRenderMultiFileMustHaveArrayAsParameter: ()=> {
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((validTemplate)=>{
                return new SpreadSheet().load(new JSZip(validTemplate));
            }).then((spreadsheet)=>{
                return spreadsheet.bulkRenderMultiFile({name:'hogehoge'});
            }).then(()=>{
                throw new Error('bulkRenderMultiFile_must_have_array_as_parameter failed ');
            }).catch((err)=>{
                assert.equal(err,'bulkRenderMultiFile() has only array object');
            });
    },

    bulkRenderMultiFileMustHaveNameAndData: ()=> {
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((validTemplate)=>{
                return new SpreadSheet().load(new JSZip(validTemplate));
            }).then((spreadsheet)=>{
                return spreadsheet.bulkRenderMultiFile([{name:'hogehoge'}]);
            }).then(()=>{
                throw new Error('bulkRenderMultiFile_must_have_name_and_data failed ');
            }).catch((err)=>{
                assert.equal(err,'bulkRenderMultiFile() is called with invalid parameter');
            });
    },

    checkIfBulkRenderMultiFileRendersCorrectly: ()=>{
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((validTemplate)=>{
                return new SpreadSheet().load(new JSZip(validTemplate));
            }).then((spreadsheet)=>{
                return spreadsheet.bulkRenderMultiFile([
                    {name:'file1.xlsx',data:{AccountName__c:'hoge account1',AccountAddress__c:'hoge street1'}},
                    {name:'file2.xlsx',data:{AccountName__c:'hoge account2',AccountAddress__c:'hoge street2'}},
                    {name:'file3.xlsx',data:{AccountName__c:'hoge account3',AccountAddress__c:'hoge street3'}}
                ]);
            }).then((zipData)=>{
                let zip = new JSZip(zipData);
                let excel1 = zip.file('file1.xlsx').asArrayBuffer();
                let excel2 = zip.file('file2.xlsx').asArrayBuffer();
                let excel3 = zip.file('file3.xlsx').asArrayBuffer();
                return Promise.props({
                    sp1: new SpreadSheet().load(new JSZip(excel1)),
                    sp2: new SpreadSheet().load(new JSZip(excel2)),
                    sp3: new SpreadSheet().load(new JSZip(excel3))
                }).then(({sp1,sp2,sp3})=>{
                    assert(sp1.hasAsSharedString('hoge account1'),"'hoge account1' is missing in excel file");
                    assert(sp1.hasAsSharedString('hoge street1'),"'hoge street1' is missing in excel file");
                    assert(sp2.hasAsSharedString('hoge account2'),"'hoge account2' is missing in excel file");
                    assert(sp2.hasAsSharedString('hoge street2'),"'hoge street2' is missing in excel file");
                    assert(sp3.hasAsSharedString('hoge account3'),"'hoge account3' is missing in excel file");
                    assert(sp3.hasAsSharedString('hoge street3'),"'hoge street3' is missing in excel file");
                });

            });
    },

    addSheetBindingDataWithNoParameterShouldReturnError: ()=>{
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((validTemplate)=>{
                return new SpreadSheet().load(new JSZip(validTemplate));
            }).then((spreadsheet)=>{
                return spreadsheet.addSheetBindingData();
            }).then(()=>{
                throw new Error('addSheetBindingData_with_no_parameter_should_return_error failed ');
            }).catch((err)=>{
                assert.equal(err,'addSheetBindingData() needs to have 2 paramter.');
            });
    },

    addSheetBindingDataWith1ParameterShouldReturnError: ()=>{
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((validTemplate)=>{
                return new SpreadSheet().load(new JSZip(validTemplate));
            }).then((spreadsheet)=>{
                return spreadsheet.addSheetBindingData('hoge');
            }).then(()=>{
                throw new Error('addSheetBindingData_with_no_parameter_should_return_error failed ');
            }).catch((err)=>{
                assert.equal(err,'addSheetBindingData() needs to have 2 paramter.');
            });
    },

    activateSheetWithNoParameterShouldReturnError: ()=>{
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((validTemplate)=>{
                return new SpreadSheet().load(new JSZip(validTemplate));
            }).then((spreadsheet)=>{
                return spreadsheet.activateSheet();
            }).then(()=>{
                throw new Error('activateSheet_with_no_parameter_should_return_error failed ');
            }).catch((err)=>{
                assert.equal(err,'activateSheet() needs to have 1 paramter.');
            });
    },

    activateSheetWithInvalidSheetnameShouldReturnError: ()=>{
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((valid_template)=>{
                return new SpreadSheet().load(new JSZip(valid_template));
            }).then((spreadsheet)=>{
                return spreadsheet.activateSheet('hoge');
            }).then(()=>{
                throw new Error('activateSheet_with_no_parameter_should_return_error failed ');
            }).catch((err)=>{
                assert.equal(err,"Invalid sheet name 'hoge'.");
            });
    },

    deleteSheetWithNoParameterShouldReturnError: ()=>{
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((valid_template)=>{
                return new SpreadSheet().load(new JSZip(valid_template));
            }).then((spreadsheet)=>{
                return spreadsheet.deleteSheet();
            }).then(()=>{
                throw new Error('deleteSheet_with_no_parameter_should_return_error failed ');
            }).catch((err)=>{
                assert.equal(err,'deleteSheet() needs to have 1 paramter.');
            });
    },

    deleteSheetWithInvalidSheetnameShouldReturnError: ()=>{
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((valid_template)=>{
                return new SpreadSheet().load(new JSZip(valid_template));
            }).then((spreadsheet)=>{
                return spreadsheet.deleteSheet('hoge');
            }).then(()=>{
                throw new Error('deleteSheet_with_invalid_sheetname_should_return_error failed ');
            }).catch((err)=>{
                assert.equal(err,"Invalid sheet name 'hoge'.");
            });
    }
};
