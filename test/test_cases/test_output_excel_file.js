/**
 * * test_output_excel_file.js
 * * Test code for spreadsheet
 * * @author Satoshi Haga
 * * @date 2015/10/11
 **/
'use strict';

var _createClass = (function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ('value' in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; })();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError('Cannot call a class as a function'); } }

var path = require('path');
var cwd = path.resolve('');
var assert = require('assert');
var JSZip = require('jszip');
var ExcelMerge = require(cwd + '/excelmerge');
var SpreadSheet = require(cwd + '/lib/spreadsheet');
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
    }]);

    return Utility;
})();

module.exports = Utility;