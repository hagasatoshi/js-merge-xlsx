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
     * * @return {Promise} Promise instance including this
     **/

    _createClass(ExcelMerge, [{
        key: 'load',
        value: function load(excel) {
            var _this = this;

            //validation
            if (!(excel instanceof JSZip)) {
                return Promise.reject('First parameter must be JSZip instance including MS-Excel data');
            }

            return this.spreadsheet.load(excel).then(function () {
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

            //validation
            if (!bindData) {
                return Promise.reject('merge() must has parameter');
            }

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

            //validation
            if (!bindDataArray) {
                return Promise.reject('bulkMergeMultiFile() must has parameter');
            }
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

            //validation
            if (!bindDataArray || !_.isArray(bindDataArray)) {
                return Promise.reject('bulkMergeMultiSheet() must has array as parameter');
            }

            _.each(bindDataArray, function (_ref) {
                var name = _ref.name;
                var data = _ref.data;
                return _this2.spreadsheet.addSheetBindingData(name, data);
            });
            return this.spreadsheet.deleteTemplateSheet().focusOnFirstSheet().generate(output_buffer);
        }
    }]);

    return ExcelMerge;
})();

module.exports = ExcelMerge;