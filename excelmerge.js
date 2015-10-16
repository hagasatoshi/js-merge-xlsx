/**
 * * ExcelMerge
 * * top level api class for js-merge-xlsx
 * * @author Satoshi Haga
 * * @date 2015/09/30
 **/

'use strict';

var _createClass = (function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ('value' in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; })();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError('Cannot call a class as a function'); } }

var Mustache = require('mustache');
var Promise = require('bluebird');
var _ = require('underscore');
var JSZip = require('jszip');
var SpreadSheet = require('./lib/spreadsheet');
var isNode = require('detect-node');
var output_buffer = { type: isNode ? 'nodebuffer' : 'blob', compression: "DEFLATE" };

var ExcelMerge = (function () {

    /**
     * * constructor
     * *
     **/

    function ExcelMerge() {
        _classCallCheck(this, ExcelMerge);

        this.spreadsheet = new SpreadSheet();
    }

    //Exports

    /**
     * * load
     * * @param {Object} excel JsZip object including MS-Excel file
     * * @param {Object} option option parameter
     * * @return {Promise} Promise instance including this
     **/

    _createClass(ExcelMerge, [{
        key: 'load',
        value: function load(excel, option) {
            var _this = this;

            return this.spreadsheet.load(excel, option).then(function () {
                return _this;
            });
        }

        /**
         * * merge
         * * @param {Object} bindData binding data
         * * @return {Promise} Promise instance including MS-Excel data. data-format is determined by jszip_option
         **/
    }, {
        key: 'merge',
        value: function merge(bindData) {
            return this.spreadsheet.simpleMerge(bindData);
        }

        /**
         * * bulkMergeMultiFile
         * * @param {Array} bindDataArray including data{name: file's name, data: binding-data}
         * * @return {Promise} Promise instance including MS-Excel data.
         **/
    }, {
        key: 'bulkMergeMultiFile',
        value: function bulkMergeMultiFile(bindDataArray) {
            return this.spreadsheet.bulkMergeMultiFile(bindDataArray);
        }

        /**
         * * bulkMergeMultiSheet
         * * @param {Array} bindDataArray including data{name: file's name, data: binding-data}
         * * @return {Promise} Promise instance including MS-Excel data.
         **/
    }, {
        key: 'bulkMergeMultiSheet',
        value: function bulkMergeMultiSheet(bindDataArray) {
            var _this2 = this;

            return bindDataArray.reduce(function (promise, _ref) {
                var name = _ref.name;
                var data = _ref.data;
                return promise.then(function (prior) {
                    return _this2.spreadsheet.addSheetBindingData(name, data);
                });
            }, Promise.resolve()).then(function () {
                return _this2.spreadsheet.deleteTemplateSheet().forcusOnFirstSheet().generate(output_buffer);
            })['catch'](function (err) {
                console.error(new Error(err).stack);
                Promise.reject();
            });
        }
    }]);

    return ExcelMerge;
})();

module.exports = ExcelMerge;