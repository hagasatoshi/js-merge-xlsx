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
        value: function output(template_name, input_file_name, output_type, output_file_name) {
            return fs.readFileAsync(__dirname + '/../templates/' + template_name).then(function (excel_template) {
                return Promise.props({
                    rendering_data: readYamlAsync(__dirname + '/../input/' + input_file_name), //Load single data
                    merge: new ExcelMerge().load(new JSZip(excel_template)) //Initialize ExcelMerge object
                });
            }).then(function (result) {
                //ExcelMerge object
                var merge = result.merge;

                //rendering data
                var rendering_data = undefined;
                if (output_type === EXCEL_OUTPUT_TYPE.SINGLE) {
                    rendering_data = result.rendering_data;
                    return merge.render(rendering_data);
                } else if (output_type === EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_FILE) {
                    rendering_data = [];
                    _.each(result.rendering_data, function (data, index) {
                        rendering_data.push({ name: 'file' + (index + 1) + '.xlsx', data: data });
                    });
                    return merge.bulkRenderMultiFile(rendering_data);
                } else if (output_type === EXCEL_OUTPUT_TYPE.BULK_MULTIPLE_SHEET) {
                    rendering_data = [];
                    _.each(result.rendering_data, function (data, index) {
                        rendering_data.push({ name: 'example' + (index + 1), data: data });
                    });
                    return merge.bulkRenderMultiSheet(rendering_data);
                }
            }).then(function (output_data) {
                return fs.writeFileAsync(__dirname + '/../output/' + output_file_name, output_data);
            }).then(function () {
                assert(true);
            })['catch'](function (err) {
                console.error(new Error(err).stack);
                assert(false);
            });
        }
    }]);

    return Utility;
})();

module.exports = Utility;