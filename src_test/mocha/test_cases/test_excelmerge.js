/**
 * test_spreadsheet.js
 * Test code for spreadsheet
 * @author Satoshi Haga
 * @date 2015/10/10
 */

const path = require('path');
const cwd = path.resolve('');
const assert = require('assert');
const JSZip = require('jszip');
const ExcelMerge = require(cwd + '/excelmerge');
const SpreadSheet = require(cwd + '/lib/sheetHelper');
require(cwd + '/lib/underscore_mixin');
const Promise = require('bluebird');
const readYamlAsync = Promise.promisify(require('read-yaml'));
const fs = Promise.promisifyAll(require('fs'));
const _ = require('underscore');

const SINGLE_DATA = 'SINGLE_DATA';
const MULTI_FILE = 'MULTI_FILE';
const MULTI_SHEET = 'MULTI_SHEET';

module.exports = {
    checkLoadWithNoParameterShouldReturnError: ()=>{
        return new ExcelMerge().load()
            .then(()=>{
                throw new Error('checkLoadWithNoParameterShouldReturnError failed ');
            }).catch((err)=>{
                assert.equal(err, 'First parameter must be JSZip instance including MS-Excel data');
            });
    },

    checkLoadShouldReturnThisInstance: ()=>{
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((validTemplate)=>{
                return new ExcelMerge().load(new JSZip(validTemplate));
            }).then((excelMerge)=>{
                assert(excelMerge instanceof ExcelMerge, 'ExcelMerge#load() should return this instance');
            });
    },

    checkVariablesWorkCorrectly: ()=>{
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((validTemplate)=>{
                return new ExcelMerge().load(new JSZip(validTemplate));
            }).then((excelMerge)=>{
                let variables = [
                    'AccountName__c', 'StartDateFormat__c', 'EndDateFormat__c', 'Address__c', 'JobDescription__c', 'StartTime__c', 'EndTime__c',
                    'hasOverTime__c', 'HoliDayType__c', 'Salary__c', 'DueDate__c', 'SalaryDate__c', 'AccountName__c', 'AccountAddress__c'
                ];
                let parsedVariables = excelMerge.variables();
                _.each(variables, (e)=>{
                    assert(_.contains(parsedVariables,e), `${e} is not parsed correctly by variables()`);
                });
            });
    },

    checkIfBulkMergeMultiSheetRendersCorrectly: ()=>{
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((validTemplate)=>{
                return new ExcelMerge().load(new JSZip(validTemplate));
            }).then((excelMerge)=>{
                return excelMerge.bulkMergeMultiSheet([
                    { name: 'sheet1', data: { AccountName__c: 'hoge account1', AccountAddress__c: 'hoge street1' } },
                    { name: 'sheet2', data: { AccountName__c: 'hoge account2', AccountAddress__c: 'hoge street2' } },
                    { name: 'sheet3', data: { AccountName__c: 'hoge account3', AccountAddress__c: 'hoge street3' } }
                ]);
            }).then((excelData)=>{
                return new SpreadSheet().load(new JSZip(excelData));
            }).then((spreadsheet)=>{
                assert(spreadsheet.hasAsSharedString('hoge account1'), "'hoge account1' is missing in excel file");
                assert(spreadsheet.hasAsSharedString('hoge street1'), "'hoge street1' is missing in excel file");
                assert(spreadsheet.hasAsSharedString('hoge account2'), "'hoge account2' is missing in excel file");
                assert(spreadsheet.hasAsSharedString('hoge street2'), "'hoge street2' is missing in excel file");
                assert(spreadsheet.hasAsSharedString('hoge account3'), "'hoge account3' is missing in excel file");
                assert(spreadsheet.hasAsSharedString('hoge street3'), "'hoge street3' is missing in excel file");
            });
    },

    checkIfMergeByTypeRendersCorrectly3: ()=>{
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((validTemplate)=>{
                return new ExcelMerge().load(new JSZip(validTemplate));
            }).then((excelMerge)=>{
                return excelMerge.mergeByType(MULTI_SHEET,[
                    { name: 'sheet1', data: { AccountName__c: 'hoge account1', AccountAddress__c: 'hoge street1' } },
                    { name: 'sheet2', data: { AccountName__c: 'hoge account2', AccountAddress__c: 'hoge street2' } },
                    { name: 'sheet3', data: { AccountName__c: 'hoge account3', AccountAddress__c: 'hoge street3' } }
                ]);
            }).then((excelData)=>{
                return new SpreadSheet().load(new JSZip(excelData));
            }).then((spreadsheet)=>{
                assert(spreadsheet.hasAsSharedString('hoge account1'), "'hoge account1' is missing in excel file");
                assert(spreadsheet.hasAsSharedString('hoge street1'), "'hoge street1' is missing in excel file");
                assert(spreadsheet.hasAsSharedString('hoge account2'), "'hoge account2' is missing in excel file");
                assert(spreadsheet.hasAsSharedString('hoge street2'), "'hoge street2' is missing in excel file");
                assert(spreadsheet.hasAsSharedString('hoge account3'), "'hoge account3' is missing in excel file");
                assert(spreadsheet.hasAsSharedString('hoge street3'), "'hoge street3' is missing in excel file");
            });
    },

    checkLoadEachMemberFromValidTemplate: ()=>{
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((validTemplate)=>{
                return new ExcelMerge().load(new JSZip(validTemplate));
            }).then((excelMerge)=>{

                //excel
                assert(excelMerge.sheetHelper.excel instanceof JSZip, 'SpreadSheet#excel is not assigned correctly');

                //check if each variables is parsed or not.
                var variables = [
                    'AccountName__c', 'StartDateFormat__c', 'EndDateFormat__c', 'JobDescription__c',
                    'StartTime__c', 'EndTime__c', 'hasOverTime__c', 'HoliDayType__c', 'Salary__c',
                    'DueDate__c', 'SalaryDate__c', 'AccountName__c', 'AccountAddress__c'
                ];
                var chkCommonStringsWithVariable = _.map(excelMerge.sheetHelper.commonStringsWithVariable, (e)=>{
                    return _.stringValue(e.t);
                });
                _.each(variables, (e)=>{
                    //variables
                    assert(_.contains(excelMerge.sheetHelper.variables, e), `ExcelMerge#load() doesn't set up ${e} as variable correctly`);
                    assert(
                        _.find(chkCommonStringsWithVariable, (v)=>{return v.indexOf(`{{${e}}}`) !== -1;}),
                        `ExcelMerge#load() doesn't set up ${e} as variable correctly`
                    );
                });
            });
    },

    checkIfMergeRendersCorrectly: ()=>{
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((validTemplate)=>{
                return new ExcelMerge().load(new JSZip(validTemplate));
            }).then((excelMerge)=>{
                return excelMerge.merge({ AccountName__c: 'hoge account', AccountAddress__c: 'hoge street' });
            }).then((excelData)=>{
                return new SpreadSheet().load(new JSZip(excelData));
            }).then(function (spreadsheet) {
                assert(spreadsheet.variables.length === 0, "ExcelMerge#merge() doesn't work correctly");
                assert(spreadsheet.hasAsSharedString('hoge account'), "'hoge account' is not rendered by SpreadSheet#simpleMerge()");
                assert(spreadsheet.hasAsSharedString('hoge street'), "'hoge street' is not rendered by SpreadSheet#simpleMerge()");
            });
    },

    checkIfMergeByTypeRendersCorrectly1: ()=>{
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((validTemplate)=>{
                return new ExcelMerge().load(new JSZip(validTemplate));
            }).then((excelMerge)=>{
                return excelMerge.mergeByType(SINGLE_DATA, { AccountName__c: 'hoge account', AccountAddress__c: 'hoge street' });
            }).then((excelData)=>{
                return new SpreadSheet().load(new JSZip(excelData));
            }).then(function (spreadsheet) {
                assert(spreadsheet.variables.length === 0, "ExcelMerge#merge() doesn't work correctly");
                assert(spreadsheet.hasAsSharedString('hoge account'), "'hoge account' is not rendered by SpreadSheet#simpleMerge()");
                assert(spreadsheet.hasAsSharedString('hoge street'), "'hoge street' is not rendered by SpreadSheet#simpleMerge()");
            });
    },
    checkIfMergeWithNoParameterRendersCorrectly: ()=>{
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((validTemplate)=>{
                return new ExcelMerge().load(new JSZip(validTemplate));
            }).then((excelMerge)=>{
                return excelMerge.merge();
            }).then(()=>{
                throw new Error('checkIfMergeWithNoParameterRendersCorrectly failed');
            }).catch((err)=>{
                assert.equal(err, 'merge() must has parameter');
            });
    },

    checkIfBulkMergeMultiFileRendersCorrectly: ()=>{
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((validTemplate)=>{
                return new ExcelMerge().load(new JSZip(validTemplate));
            }).then((excelMerge)=>{
                return excelMerge.bulkMergeMultiFile([
                    { name: 'file1.xlsx', data: { AccountName__c: 'hoge account1', AccountAddress__c: 'hoge street1' } },
                    { name: 'file2.xlsx', data: { AccountName__c: 'hoge account2', AccountAddress__c: 'hoge street2' } },
                    { name: 'file3.xlsx', data: { AccountName__c: 'hoge account3', AccountAddress__c: 'hoge street3' } }
                ]);
            }).then((zipData)=>{
                var zip = new JSZip(zipData);
                var excel1 = zip.file('file1.xlsx').asArrayBuffer();
                var excel2 = zip.file('file2.xlsx').asArrayBuffer();
                var excel3 = zip.file('file3.xlsx').asArrayBuffer();
                return Promise.props({
                    sp1: new SpreadSheet().load(new JSZip(excel1)),
                    sp2: new SpreadSheet().load(new JSZip(excel2)),
                    sp3: new SpreadSheet().load(new JSZip(excel3))
                }).then(({sp1,sp2,sp3})=>{

                    assert(sp1.hasAsSharedString('hoge account1'), "'hoge account1' is missing in excel file");
                    assert(sp1.hasAsSharedString('hoge street1'), "'hoge street1' is missing in excel file");
                    assert(sp2.hasAsSharedString('hoge account2'), "'hoge account2' is missing in excel file");
                    assert(sp2.hasAsSharedString('hoge street2'), "'hoge street2' is missing in excel file");
                    assert(sp3.hasAsSharedString('hoge account3'), "'hoge account3' is missing in excel file");
                    assert(sp3.hasAsSharedString('hoge street3'), "'hoge street3' is missing in excel file");
                });
            });
    },
    checkIfMergeByTypeRendersCorrectly2: ()=>{
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((validTemplate)=>{
                return new ExcelMerge().load(new JSZip(validTemplate));
            }).then((excelMerge)=>{
                return excelMerge.mergeByType(MULTI_FILE, [
                    { name: 'file1.xlsx', data: { AccountName__c: 'hoge account1', AccountAddress__c: 'hoge street1' } },
                    { name: 'file2.xlsx', data: { AccountName__c: 'hoge account2', AccountAddress__c: 'hoge street2' } },
                    { name: 'file3.xlsx', data: { AccountName__c: 'hoge account3', AccountAddress__c: 'hoge street3' } }
                ]);
            }).then((zipData)=>{
                var zip = new JSZip(zipData);
                var excel1 = zip.file('file1.xlsx').asArrayBuffer();
                var excel2 = zip.file('file2.xlsx').asArrayBuffer();
                var excel3 = zip.file('file3.xlsx').asArrayBuffer();
                return Promise.props({
                    sp1: new SpreadSheet().load(new JSZip(excel1)),
                    sp2: new SpreadSheet().load(new JSZip(excel2)),
                    sp3: new SpreadSheet().load(new JSZip(excel3))
                }).then(({sp1,sp2,sp3})=>{

                    assert(sp1.hasAsSharedString('hoge account1'), "'hoge account1' is missing in excel file");
                    assert(sp1.hasAsSharedString('hoge street1'), "'hoge street1' is missing in excel file");
                    assert(sp2.hasAsSharedString('hoge account2'), "'hoge account2' is missing in excel file");
                    assert(sp2.hasAsSharedString('hoge street2'), "'hoge street2' is missing in excel file");
                    assert(sp3.hasAsSharedString('hoge account3'), "'hoge account3' is missing in excel file");
                    assert(sp3.hasAsSharedString('hoge street3'), "'hoge street3' is missing in excel file");
                });
            });
    },
    checkIfBulkMergeMultiFileWithNoParameterShouldReturnError: ()=>{
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((validTemplate)=>{
                return new ExcelMerge().load(new JSZip(validTemplate));
            }).then((excelMerge)=>{
                return excelMerge.bulkMergeMultiFile();
            }).then(()=>{
                throw new Error('checkIfBulkMergeMultiFileWithNoParameterShouldReturnError failed');
            }).catch((err)=>{
                assert.equal(err, 'bulkMergeMultiFile() must has parameter');
            });
    },

    checkIfBulkMergeMultiSheetWithNoParameterShouldReturnError: ()=>{
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((validTemplate)=>{
                return new ExcelMerge().load(new JSZip(validTemplate));
            }).then((excelMerge)=>{
                return excelMerge.bulkMergeMultiSheet();
            }).then(()=>{
                throw new Error('checkIfBulkMergeMultiSheetWithNoParameterShouldReturnError failed');
            }).catch((err)=>{
                assert.equal(err, 'bulkMergeMultiSheet() must has array as parameter');
            });
    },

    checkIfMergeByTypeThrowErrorWithInvalidType: ()=>{
        return fs.readFileAsync(__dirname + '/../templates/Template.xlsx')
            .then((validTemplate)=>{
                return new ExcelMerge().load(new JSZip(validTemplate));
            }).then((excelMerge)=>{
                return excelMerge.mergeByType('hoge', [
                    { name: 'file1.xlsx', data: { AccountName__c: 'hoge account1', AccountAddress__c: 'hoge street1' } },
                    { name: 'file2.xlsx', data: { AccountName__c: 'hoge account2', AccountAddress__c: 'hoge street2' } },
                    { name: 'file3.xlsx', data: { AccountName__c: 'hoge account3', AccountAddress__c: 'hoge street3' } }
                ]);
            }).then(()=>{
                throw new Error('checkIfMergeByTypeThrowErrorWithInvalidType failed');
            }).catch((err)=>{
                assert.equal(err, 'Invalid parameter : mergeType');
            });
    }

};