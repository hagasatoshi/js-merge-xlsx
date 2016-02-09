/**
 * test_output_excel_file.js
 * Test code for spreadsheet
 * @author Satoshi Haga
 * @date 2015/10/11
 */
'use strict';

var _createClass = (function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ('value' in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; })();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError('Cannot call a class as a function'); } }

var path = require('path');
var cwd = path.resolve('');
var assert = require('assert');
var JSZip = require('jszip');
var ExcelMerge = require(cwd + '/excelmerge');
var SpreadSheet = require(cwd + '/lib/sheetHelper');
require(cwd + '/lib/underscore_mixin');
var Promise = require('bluebird');
var readYamlAsync = Promise.promisify(require('read-yaml'));
var fs = Promise.promisifyAll(require('fs'));
var _ = require('underscore');

var EXCEL_OUTPUT_TYPE = {
    SINGLE: 0,
    BULK_MULTIPLE_FILE: 1,
    BULK_MULTIPLE_SHEET: 2
};

var Utility = (function () {
    function Utility() {
        _classCallCheck(this, Utility);
    }

    _createClass(Utility, [{
        key: 'output',
        value: function output(templateName, inputFileName, outputType, outputFileName) {
            return fs.readFileAsync(__dirname + '/../templates/' + templateName).then(function (excelTemplate) {
                return Promise.props({
                    renderingData: readYamlAsync(__dirname + '/../input/' + inputFileName), //Load single data
                    excelMerge: new ExcelMerge().load(new JSZip(excelTemplate)) //Initialize ExcelMerge object
                });
            }).then(function (_ref) {
                var renderingData = _ref.renderingData;
                var excelMerge = _ref.excelMerge;

                var dataArray = [];
                switch (outputType) {
                    case EXCEL_OUTPUT_TYPE.SINGLE:
                        return excelMerge.merge(renderingData);
                        break;

                    case EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_FILE:
                        _.each(renderingData, function (data, index) {
                            return dataArray.push({ name: 'file' + (index + 1) + '.xlsx', data: data });
                        });
                        return excelMerge.bulkMergeMultiFile(dataArray);
                        break;

                    case EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_SHEET:
                        _.each(renderingData, function (data, index) {
                            return dataArray.push({ name: 'example' + (index + 1), data: data });
                        });
                        return excelMerge.bulkMergeMultiSheet(dataArray);
                        break;
                }
            }).then(function (outputData) {
                return fs.writeFileAsync(__dirname + '/../output/' + outputFileName, outputData);
            }).then(function () {
                return assert(true);
            })['catch'](function (err) {
                console.error(new Error(err).stack);
                assert(false);
            });
        }
    }, {
        key: 'output_character_test_single_record',
        value: function output_character_test_single_record(templateName, outputFileName) {
            return fs.readFileAsync(__dirname + '/../templates/' + templateName).then(function (excelTemplate) {
                return new ExcelMerge().load(new JSZip(excelTemplate));
            }).then(function (excelMerge) {
                var renderingData = {
                    AccountName__c: '<>"\'&\'(0=0|~|==0~==0)=((\'(\'&\'%%&%%\'$%$',
                    AccountAddress__c: "KSPOI0)I0I0K0)(()')('#)JOKJ_?><<MNNBVCXXZ",
                    StartDateFormat__c: '2015/10/01',
                    EndDateFormat__c: '2016-9-30',
                    Address__c: '！イ”＝０｀M＝０｀イ＝『＝０『オ＝〜＝オ＝〜KW＝｀イ）＝｀０）！｜『？ァQ',
                    JobDescription__c: '①②③④⑤',
                    StartTime__c: '@@@@@@',
                    EndTime__c: '奧應橫歐毆穩假價畫會囘懷繪擴殼覺學嶽樂勸嵩'
                };
                return excelMerge.merge(renderingData);
            }).then(function (outputData) {
                return fs.writeFileAsync(__dirname + '/../output/' + outputFileName, outputData);
            }).then(function () {
                return assert(true);
            })['catch'](function (err) {
                console.error(new Error(err).stack);
                assert(false);
            });
        }
    }, {
        key: 'output_character_test_bulk_record_as_multifile',
        value: function output_character_test_bulk_record_as_multifile(templateName, outputFileName) {
            return fs.readFileAsync(__dirname + '/../templates/' + templateName).then(function (excelTemplate) {
                return new ExcelMerge().load(new JSZip(excelTemplate));
            }).then(function (excelMerge) {
                var renderingData = [{
                    name: 'file1.xlsx',
                    data: {
                        AccountName__c: '<>"\'&\'(0=0|~|==0~==0)=((\'(\'&\'%%&%%\'$%$',
                        AccountAddress__c: "KSPOI0)I0I0K0)(()')('#)JOKJ_?><<MNNBVCXXZ",
                        StartDateFormat__c: '2015/10/01',
                        EndDateFormat__c: '2016-9-30',
                        Address__c: '！イ”＝０｀M＝０｀イ＝『＝０『オ＝〜＝オ＝〜KW＝｀イ）＝｀０）！｜『？ァQ',
                        JobDescription__c: '①②③④⑤',
                        StartTime__c: '@@@@@@',
                        EndTime__c: '奧應橫歐毆穩假價畫會囘懷繪擴殼覺學嶽樂勸嵩'
                    }
                }, {
                    name: 'file2.xlsx',
                    data: {
                        AccountName__c: '<>"\'&\'(0=0|~|==0~==0)=((\'(\'&\'%%&%%\'$%$',
                        AccountAddress__c: "KSPOI0)I0I0K0)(()')('#)JOKJ_?><<MNNBVCXXZ",
                        StartDateFormat__c: '2015/10/01',
                        EndDateFormat__c: '2016-9-30',
                        Address__c: '！イ”＝０｀M＝０｀イ＝『＝０『オ＝〜＝オ＝〜KW＝｀イ）＝｀０）！｜『？ァQ',
                        JobDescription__c: '①②③④⑤',
                        StartTime__c: '@@@@@@',
                        EndTime__c: '奧應橫歐毆穩假價畫會囘懷繪擴殼覺學嶽樂勸嵩'
                    }
                }, {
                    name: 'file3.xlsx',
                    data: {
                        AccountName__c: '<>"\'&\'(0=0|~|==0~==0)=((\'(\'&\'%%&%%\'$%$',
                        AccountAddress__c: "KSPOI0)I0I0K0)(()')('#)JOKJ_?><<MNNBVCXXZ",
                        StartDateFormat__c: '2015/10/01',
                        EndDateFormat__c: '2016-9-30',
                        Address__c: '！イ”＝０｀M＝０｀イ＝『＝０『オ＝〜＝オ＝〜KW＝｀イ）＝｀０）！｜『？ァQ',
                        JobDescription__c: '①②③④⑤',
                        StartTime__c: '@@@@@@',
                        EndTime__c: '奧應橫歐毆穩假價畫會囘懷繪擴殼覺學嶽樂勸嵩'
                    }
                }];
                return excelMerge.bulkMergeMultiFile(renderingData);
            }).then(function (outputData) {
                return fs.writeFileAsync(__dirname + '/../output/' + outputFileName, outputData);
            }).then(function () {
                return assert(true);
            })['catch'](function (err) {
                console.error(new Error(err).stack);
                assert(false);
            });
        }
    }, {
        key: 'output_character_test_bulk_record_as_multisheet',
        value: function output_character_test_bulk_record_as_multisheet(templateName, outputFileName) {
            return fs.readFileAsync(__dirname + '/../templates/' + templateName).then(function (excelTemplate) {
                return new ExcelMerge().load(new JSZip(excelTemplate));
            }).then(function (excelMerge) {
                var renderingData = [{
                    name: 'sheet1',
                    data: {
                        AccountName__c: '<>"\'&\'(0=0|~|==0~==0)=((\'(\'&\'%%&%%\'$%$',
                        AccountAddress__c: "KSPOI0)I0I0K0)(()')('#)JOKJ_?><<MNNBVCXXZ",
                        StartDateFormat__c: '2015/10/01',
                        EndDateFormat__c: '2016-9-30',
                        Address__c: '！イ”＝０｀M＝０｀イ＝『＝０『オ＝〜＝オ＝〜KW＝｀イ）＝｀０）！｜『？ァQ',
                        JobDescription__c: '①②③④⑤',
                        StartTime__c: '@@@@@@',
                        EndTime__c: '奧應橫歐毆穩假價畫會囘懷繪擴殼覺學嶽樂勸嵩'
                    }
                }, {
                    name: 'sheet2',
                    data: {
                        AccountName__c: '<>"\'&\'(0=0|~|==0~==0)=((\'(\'&\'%%&%%\'$%$',
                        AccountAddress__c: "KSPOI0)I0I0K0)(()')('#)JOKJ_?><<MNNBVCXXZ",
                        StartDateFormat__c: '2015/10/01',
                        EndDateFormat__c: '2016-9-30',
                        Address__c: '！イ”＝０｀M＝０｀イ＝『＝０『オ＝〜＝オ＝〜KW＝｀イ）＝｀０）！｜『？ァQ',
                        JobDescription__c: '①②③④⑤',
                        StartTime__c: '@@@@@@',
                        EndTime__c: '奧應橫歐毆穩假價畫會囘懷繪擴殼覺學嶽樂勸嵩'
                    }
                }, {
                    name: 'sheet3',
                    data: {
                        AccountName__c: '<>"\'&\'(0=0|~|==0~==0)=((\'(\'&\'%%&%%\'$%$',
                        AccountAddress__c: "KSPOI0)I0I0K0)(()')('#)JOKJ_?><<MNNBVCXXZ",
                        StartDateFormat__c: '2015/10/01',
                        EndDateFormat__c: '2016-9-30',
                        Address__c: '！イ”＝０｀M＝０｀イ＝『＝０『オ＝〜＝オ＝〜KW＝｀イ）＝｀０）！｜『？ァQ',
                        JobDescription__c: '①②③④⑤',
                        StartTime__c: '@@@@@@',
                        EndTime__c: '奧應橫歐毆穩假價畫會囘懷繪擴殼覺學嶽樂勸嵩'
                    }
                }];
                return excelMerge.bulkMergeMultiSheet(renderingData);
            }).then(function (outputData) {
                return fs.writeFileAsync(__dirname + '/../output/' + outputFileName, outputData);
            }).then(function () {
                return assert(true);
            })['catch'](function (err) {
                console.error(new Error(err).stack);
                assert(false);
            });
        }
    }]);

    return Utility;
})();

module.exports = Utility;