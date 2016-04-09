/**
 * test_spreadsheet.js
 * Test code for spreadsheet
 * @author Satoshi Haga
 * @date 2015/10/10
 */
var path = require('path');
var cwd = path.resolve('');
var assert = require('assert');
const Excel = require(cwd + '/lib/Excel');
var SpreadSheet = require(cwd + '/lib/sheetHelper');
require(cwd+'/lib/underscore_mixin');
var Promise = require('bluebird');
var readYamlAsync = Promise.promisify(require('read-yaml'));
var fs = Promise.promisifyAll(require('fs'));
var _ = require('underscore');
var isNode = require('detect-node');
const output_buffer = {type: (isNode?'nodebuffer':'blob'), compression:"DEFLATE"};

module.exports = {

    checkLoadShouldReturnThisInstance: ()=>{
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((validTemplate)=>{
                return new SpreadSheet().load(new Excel(validTemplate));
            }).then((spreadsheet)=>{
                assert(spreadsheet instanceof SpreadSheet, 'SpreadSheet#load() should return this instance');
            });
    },

    checkLoadEachMemberFromValidTemplate: ()=>{
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((validTemplate)=>{
                return new SpreadSheet().load(new Excel(validTemplate));
            }).then((spreadsheet)=>{

                //excel
                assert(spreadsheet.excel instanceof Excel, 'SpreadSheet#excel is not assigned correctly');

                //check if each variables is parsed or not.
                let variables = [
                    'AccountName__c','StartDateFormat__c','EndDateFormat__c','JobDescription__c','StartTime__c',
                    'EndTime__c','hasOverTime__c','HoliDayType__c','Salary__c','DueDate__c','SalaryDate__c',
                    'AccountName__c','AccountAddress__c'
                ];
                var chkCommonStringsWithVariable = _.map(spreadsheet.commonStringsWithVariable,(e)=>_(e.t).stringValue());
                _.each(variables, (e)=>{
                    //variables
                    assert(_.contains(spreadsheet.excel.variables(), e), `SpreadSheet#load() doesn't set up ${e} as variable correctly`);
                    assert(_.find(chkCommonStringsWithVariable, (v)=>(v.indexOf(`{{${e}}}`) !== -1)), `SpreadSheet#load() doesn't set up ${e} as variable correctly`);
                });

            });
    },

    checkTemplateVariablesWorkCorrectly: ()=>{
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((validTemplate)=>{
                return new SpreadSheet().load(new Excel(validTemplate));
            }).then((spreadsheet)=>{
                let variables = [
                    'AccountName__c', 'StartDateFormat__c', 'EndDateFormat__c', 'Address__c', 'JobDescription__c', 'StartTime__c', 'EndTime__c',
                    'hasOverTime__c', 'HoliDayType__c', 'Salary__c', 'DueDate__c', 'SalaryDate__c', 'AccountName__c', 'AccountAddress__c'
                ];
                let parsedVariables = spreadsheet.excel.variables();
                _.each(variables, (e)=>{
                    assert(_.contains(parsedVariables,e), `${e} is not parsed correctly by variables()`);
                });
            });
    },

    checkIfSimpleMergeRendersCorrectly: ()=>{
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((validTemplate)=>{
                return new SpreadSheet().load(new Excel(validTemplate));
            }).then((spreadsheet)=>{
                return spreadsheet.simpleMerge({AccountName__c:'hoge account',AccountAddress__c:'hoge street'});
            }).then((excelData)=>{
                return new SpreadSheet().load(new Excel(excelData));
            }).then((spreadsheet)=>{
                assert(spreadsheet.excel.variables().length === 0, "SpreadSheet#simpleMerge() doesn't work correctly");
                assert(spreadsheet.excel.hasAsSharedString('hoge account'), "'hoge account' is not rendered by SpreadSheet#simpleMerge()");
                assert(spreadsheet.excel.hasAsSharedString('hoge street'), "'hoge street' is not rendered by SpreadSheet#simpleMerge()");
            });
    },

    checkIfBulkMergeMultiFileRendersCorrectly: ()=>{
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((validTemplate)=>{
                return new SpreadSheet().load(new Excel(validTemplate));
            }).then((spreadsheet)=>{
                return spreadsheet.bulkMergeMultiFile([
                    {name:'file1.xlsx',data:{AccountName__c:'hoge account1',AccountAddress__c:'hoge street1'}},
                    {name:'file2.xlsx',data:{AccountName__c:'hoge account2',AccountAddress__c:'hoge street2'}},
                    {name:'file3.xlsx',data:{AccountName__c:'hoge account3',AccountAddress__c:'hoge street3'}}
                ]);
            }).then((zipData)=>{
                let zip = new Excel(zipData);
                let excel1 = zip.file('file1.xlsx').asArrayBuffer();
                let excel2 = zip.file('file2.xlsx').asArrayBuffer();
                let excel3 = zip.file('file3.xlsx').asArrayBuffer();
                return Promise.props({
                    sp1: new SpreadSheet().load(new Excel(excel1)),
                    sp2: new SpreadSheet().load(new Excel(excel2)),
                    sp3: new SpreadSheet().load(new Excel(excel3))
                }).then(({sp1,sp2,sp3})=>{
                    assert(sp1.excel.hasAsSharedString('hoge account1'),"'hoge account1' is missing in excel file");
                    assert(sp1.excel.hasAsSharedString('hoge street1'),"'hoge street1' is missing in excel file");
                    assert(sp2.excel.hasAsSharedString('hoge account2'),"'hoge account2' is missing in excel file");
                    assert(sp2.excel.hasAsSharedString('hoge street2'),"'hoge street2' is missing in excel file");
                    assert(sp3.excel.hasAsSharedString('hoge account3'),"'hoge account3' is missing in excel file");
                    assert(sp3.excel.hasAsSharedString('hoge street3'),"'hoge street3' is missing in excel file");
                });

            });
    },

    checkIfBulkMergeMultiSheetRendersCorrectly: ()=>{
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((validTemplate)=>{
                return new SpreadSheet().load(new Excel(validTemplate));
            }).then((spreadsheet)=>{
                return spreadsheet.bulkMergeMultiSheet([
                    {name:'sheet1',data:{AccountName__c:'hoge account1',AccountAddress__c:'hoge street1'}},
                    {name:'sheet2',data:{AccountName__c:'hoge account2',AccountAddress__c:'hoge street2'}},
                    {name:'sheet3',data:{AccountName__c:'hoge account3',AccountAddress__c:'hoge street3'}}
                ]);
            }).then((excelData)=>{
                return new SpreadSheet().load(new Excel(excelData));
            }).then((spreadsheet)=>{
                assert(spreadsheet.excel.hasAsSharedString('hoge account1'),"'hoge account1' is missing in excel file");
                assert(spreadsheet.excel.hasAsSharedString('hoge street1'),"'hoge street1' is missing in excel file");
                assert(spreadsheet.excel.hasAsSharedString('hoge account2'),"'hoge account2' is missing in excel file");
                assert(spreadsheet.excel.hasAsSharedString('hoge street2'),"'hoge street2' is missing in excel file");
                assert(spreadsheet.excel.hasAsSharedString('hoge account3'),"'hoge account3' is missing in excel file");
                assert(spreadsheet.excel.hasAsSharedString('hoge street3'),"'hoge street3' is missing in excel file");
            });
    },

    checkIfAddSheetBindingDataCorrectly: ()=>{
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((validTemplate)=>{
                return new SpreadSheet().load(new Excel(validTemplate));
            }).then((spreadsheet)=>{
                return spreadsheet
                    .addSheetBindingData('sample', {AccountName__c:'hoge account1',AccountAddress__c:'hoge street1'})
                    .generate(output_buffer);
            }).then((excelData)=>{
                return new SpreadSheet().load(new Excel(excelData));
            }).then((spreadsheet)=>{
                assert(spreadsheet.excel.hasAsSharedString('hoge account1'),"'hoge account1' is missing in excel file");
                assert(spreadsheet.excel.hasAsSharedString('hoge street1'),"'hoge street1' is missing in excel file");
            });
    },

    checkIfDeleteTemplateSheetWorksCorrectly: ()=>{
        return fs.readFileAsync(`${__dirname}/../templates/Template.xlsx`)
            .then((validTemplate)=>{
                return new SpreadSheet().load(new Excel(validTemplate));
            }).then((spreadsheet)=>{
                return spreadsheet
                    .addSheetBindingData('sample', {AccountName__c:'hoge account1',AccountAddress__c:'hoge street1'})
                    .generate(output_buffer);
            }).then((excelData)=>{
                return new SpreadSheet().load(new Excel(excelData));
            }).then((spreadsheet)=>{
                assert(!spreadsheet.hasSheet('Sheet1'),"deleteTemplateSheet() doesn't work correctly");
                assert(spreadsheet.hasSheet('sample'),"deleteTemplateSheet() doesn't work correctly");
            });
    }
};
