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
var ExcelMerge = require(cwd+'/excelmerge');
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
                return new ExcelMerge().load(new JSZip(validTemplate));
            }).then((excelMerge)=>{
                assert(excelMerge instanceof ExcelMerge, 'ExcelMerge#load() should return this instance');
            });
    },

    checkIfBulkMergeMultiSheetRendersCorrectly: ()=>{
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((validTemplate)=>{
                return new ExcelMerge().load(new JSZip(validTemplate));
            }).then((excelMerge)=>{
                return excelMerge.bulkMergeMultiSheet([
                    {name:'sheet1',data:{AccountName__c:'hoge account1',AccountAddress__c:'hoge street1'}},
                    {name:'sheet2',data:{AccountName__c:'hoge account2',AccountAddress__c:'hoge street2'}},
                    {name:'sheet3',data:{AccountName__c:'hoge account3',AccountAddress__c:'hoge street3'}}
                ]);
            }).then((excelData)=>{
                return new SpreadSheet().load(new JSZip(excelData));
            }).then((spreadsheet)=>{
                assert(spreadsheet.hasAsSharedString('hoge account1'),"'hoge account1' is missing in excel file");
                assert(spreadsheet.hasAsSharedString('hoge street1'),"'hoge street1' is missing in excel file");
                assert(spreadsheet.hasAsSharedString('hoge account2'),"'hoge account2' is missing in excel file");
                assert(spreadsheet.hasAsSharedString('hoge street2'),"'hoge street2' is missing in excel file");
                assert(spreadsheet.hasAsSharedString('hoge account3'),"'hoge account3' is missing in excel file");
                assert(spreadsheet.hasAsSharedString('hoge street3'),"'hoge street3' is missing in excel file");
            });
    },

    checkLoadEachMemberFromValidTemplate: ()=>{
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((validTemplate)=>{
                return new ExcelMerge().load(new JSZip(validTemplate));
            }).then((excelMerge)=>{

                //excel
                assert(excelMerge.spreadsheet.excel instanceof JSZip, 'SpreadSheet#excel is not assigned correctly');

                //check if each variables is parsed or not.
                let variables = [
                    'AccountName__c','StartDateFormat__c','EndDateFormat__c','JobDescription__c','StartTime__c',
                    'EndTime__c','hasOverTime__c','HoliDayType__c','Salary__c','DueDate__c','SalaryDate__c',
                    'AccountName__c','AccountAddress__c'
                ];
                var chkCommonStringsWithVariable = _.map(excelMerge.spreadsheet.commonStringsWithVariable,(e)=>_(e.t).stringValue());
                _.each(variables, (e)=>{
                    //variables
                    assert(_.contains(excelMerge.spreadsheet.variables, e), `ExcelMerge#load() doesn't set up ${e} as variable correctly`);
                    assert(_.find(chkCommonStringsWithVariable, (v)=>(v.indexOf(`{{${e}}}`) !== -1)), `ExcelMerge#load() doesn't set up ${e} as variable correctly`);
                });

            });
    },

    checkIfMergeRendersCorrectly: ()=>{
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((validTemplate)=>{
                return new ExcelMerge().load(new JSZip(validTemplate));
            }).then((excelMerge)=>{
                return excelMerge.merge({AccountName__c:'hoge account',AccountAddress__c:'hoge street'});
            }).then((excelData)=>{
                return new SpreadSheet().load(new JSZip(excelData));
            }).then((spreadsheet)=>{
                assert(spreadsheet.variables.length === 0, "ExcelMerge#merge() doesn't work correctly");
                assert(spreadsheet.hasAsSharedString('hoge account'), "'hoge account' is not rendered by SpreadSheet#simpleMerge()");
                assert(spreadsheet.hasAsSharedString('hoge street'), "'hoge street' is not rendered by SpreadSheet#simpleMerge()");
            });
    },

    checkIfMergeWithNoParameterRendersCorrectly: ()=>{
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((validTemplate)=>{
                return new ExcelMerge().load(new JSZip(validTemplate));
            }).then((excelMerge)=>{
                return excelMerge.merge();
            }).then(()=>{
                throw new Error('checkIfMergeWithNoParameterRendersCorrectly failed');
            }).catch((err)=>{
                assert.equal(err.message,'merge() must has parameter');
            });
    },

    checkIfBulkMergeMultiFileRendersCorrectly: ()=>{
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((validTemplate)=>{
                return new ExcelMerge().load(new JSZip(validTemplate));
            }).then((excelMerge)=>{
                return excelMerge.bulkMergeMultiFile([
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

    checkIfBulkMergeMultiFileWithNoParameterShouldReturnError: ()=>{
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((validTemplate)=>{
                return new ExcelMerge().load(new JSZip(validTemplate));
            }).then((excelMerge)=>{
                return excelMerge.bulkMergeMultiFile();
            }).then(()=>{
                throw new Error('checkIfBulkMergeMultiFileWithNoParameterShouldReturnError failed');
            }).catch((err)=>{
                assert.equal(err.message,'bulkMergeMultiFile() must has parameter');
            });
    },

    checkIfBulkMergeMultiSheetWithNoParameterShouldReturnError: ()=>{
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((validTemplate)=>{
                return new ExcelMerge().load(new JSZip(validTemplate));
            }).then((excelMerge)=>{
                return excelMerge.bulkMergeMultiSheet();
            }).then(()=>{
                throw new Error('checkIfBulkMergeMultiSheetWithNoParameterShouldReturnError failed');
            }).catch((err)=>{
                assert.equal(err.message,'bulkMergeMultiSheet() must has array as parameter');
            });
    }

};