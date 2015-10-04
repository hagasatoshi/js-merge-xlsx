/**
 * * SpreadSheet
 * * Manage MS-Excel file
 * * @author Satoshi Haga
 * * @date 2015/10/03
 **/
'use strict';

var _createClass = (function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ('value' in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; })();

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { 'default': obj }; }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError('Cannot call a class as a function'); } }

var _mustache = require('mustache');

var _mustache2 = _interopRequireDefault(_mustache);

var _bluebird = require('bluebird');

var _bluebird2 = _interopRequireDefault(_bluebird);

var _underscore = require('underscore');

var _underscore2 = _interopRequireDefault(_underscore);

require('./underscore_mixin');

var _jszip = require('jszip');

var _jszip2 = _interopRequireDefault(_jszip);

var _xml2js = require('xml2js');

var _xml2js2 = _interopRequireDefault(_xml2js);

var parseString = _bluebird2['default'].promisify(_xml2js2['default'].parseString);
var builder = new _xml2js2['default'].Builder();

var SpreadSheet = (function () {
    function SpreadSheet() {
        _classCallCheck(this, SpreadSheet);
    }

    //Exports

    _createClass(SpreadSheet, [{
        key: 'load',

        /** member variables */
        //excel : {Object} JSZip instance including template excel file
        //variables : {Array} including mustache-variables defined in sharedstrings.xml
        //sharedstrings : {Array} includings common strings defined in sharedstrings.xml
        //sharedstrings_obj : {Object} whole sharedstrings object
        //sharedstrings_str : {String} whole sharedstrings string
        //common_strings_with_variable : {Array} including common strings only having mustache variables
        //sheet_xmls : {Array} including objects parsed from  'xl/worksheets/*.xml'
        //template_sheet_data : {Object} object parsed from 'xl/worksheets/*.xml'. this is used as template-file
        //template_sheet_name : {String} sheet-name of template-file
        //workbookxml_rels : {Object} parsed from 'xl/_rels/workbook.xml.rels'
        //workbookxml : {Object} parsed from 'xl/workbook.xml'

        /**
         * * load
         * * @param {Object} excel JsZip object including MS-Excel file
         * * @return {Promise|Object} Promise instance including this
         **/
        value: function load(excel) {
            var _this = this;

            //set member variable
            this.excel = excel;
            this.sharedstrings_str = excel.file('xl/sharedStrings.xml').asText();
            this.variables = (0, _underscore2['default'])(this.sharedstrings_str).variables();
            this.common_strings_with_variable = [];

            //some members are parsed in promise-chain because xml2js parses asynchronously
            return _bluebird2['default'].props({
                sharedstrings_obj: parseString(this.sharedstrings_str),
                workbookxml_rels: parseString(this.excel.file('xl/_rels/workbook.xml.rels').asText()),
                workbookxml: parseString(this.excel.file('xl/workbook.xml').asText()),
                sheet_xmls: this._parse_dir_in_excel('xl/worksheets')
            }).then(function (templates) {
                _this.sharedstrings_obj = templates.sharedstrings_obj;
                _this.sharedstrings = templates.sharedstrings_obj.sst.si;
                _this.workbookxml_rels = templates.workbookxml_rels;
                _this.workbookxml = templates.workbookxml;
                _this.sheet_xmls = templates.sheet_xmls;
                _this.template_sheet_data = _underscore2['default'].find(templates.sheet_xmls, function (e) {
                    return e.name.indexOf('.rels') === -1;
                }).worksheet.sheetData[0].row;
                _this.template_sheet_name = _this.workbookxml.workbook.sheets[0].sheet[0]['$'].name;
                _this.common_strings_with_variable = _this._parse_common_string_with_variable();

                //return this for chaining
                return _this;
            });
        }

        /**
         * * simple_render
         * * @param {Object} bind_data binding data
         * * @param {Object} jszip_option JsZip#generate() option.
         * * @returns {Promise|Object} rendered MS-Excel data. data-format is determined by jszip_option
         **/
    }, {
        key: 'simple_render',
        value: function simple_render(bind_data, jszip_option) {
            var _this2 = this;

            return _bluebird2['default'].resolve().then(function () {
                return _this2._simple_render(bind_data, jszip_option);
            });
        }

        /**
         * * bulk_render_multi_file
         * * @param {Array} bind_data_array including data{name: file's name, data: binding-data}
         * * @param {Object} jszip_option JsZip#generate() option.
         * * @returns {Promise|Object} rendered MS-Excel data.
         **/
    }, {
        key: 'bulk_render_multi_file',
        value: function bulk_render_multi_file(bind_data_array, jszip_option) {
            var _this3 = this;

            var all_excels = new _jszip2['default']();
            _underscore2['default'].each(bind_data_array, function (bind_data) {
                all_excels.file(bind_data.name, _this3._simple_render(bind_data.data, jszip_option));
            });
            return _bluebird2['default'].resolve().then(function () {
                return all_excels.generate(jszip_option);
            });
        }

        /**
         * * add_sheet_binding_data
         * * @param {String} dest_sheet_name name of new sheet
         * * @param {Object} data binding data
         * * @return {Promise|Object} Excel data. format is determinated by parameter
         **/
    }, {
        key: 'add_sheet_binding_data',
        value: function add_sheet_binding_data(dest_sheet_name, data) {
            var _this4 = this;

            //1.add relation of next sheet
            var next_id = this._available_sheetid();
            this.workbookxml_rels.Relationships.Relationship.push({ '$': { Id: next_id,
                    Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                    Target: 'worksheets/sheet' + next_id + '.xml'
                }
            });
            this.workbookxml.workbook.sheets[0].sheet.push({ '$': { name: dest_sheet_name, sheetId: next_id.replace('rId', ''), 'r:id': next_id } });

            //2.add sheet file.
            //2-1.prepare rendered-strings
            var rendered_strings = JSON.parse(JSON.stringify(this.common_strings_with_variable));
            _underscore2['default'].each(rendered_strings, function (e) {
                e.t[0] = _mustache2['default'].render((0, _underscore2['default'])(e.t).string_value(), data);
            });

            //2-2.add rendered-string into sharedstrings
            var current_count = this.sharedstrings.length;
            _underscore2['default'].each(rendered_strings, function (e, index) {
                e.shared_index = current_count + index;
                _this4.sharedstrings.push(e);
            });

            //2-4.build new sheet oject
            var source_sheet = this._sheet_by_name(this.template_sheet_name).value;
            var added_sheet = this._build_new_sheet(source_sheet, rendered_strings);

            //2-5.update sheet name.
            added_sheet.name = 'sheet' + next_id + '.xml';

            //2-6.add this sheet into sheet_xmls
            this.sheet_xmls.push(added_sheet);
        }

        /**
        * * generate
        * * call JSZip#generate() binding current data
        * * @param {Object} option option for JsZip#genereate()
        * * @return {Object} Excel data. format is determinated by parameter
        **/
    }, {
        key: 'generate',
        value: function generate(option) {
            var _this5 = this;

            //sharedstrings
            this.sharedstrings_obj.sst.si = this.sharedstrings;
            this.sharedstrings_obj.sst['$'].count = this.sharedstrings_obj.sst['$'].uniqueCount = this.sharedstrings.length;
            this.excel.file('xl/sharedStrings.xml', builder.buildObject(this.sharedstrings_obj));
            //workbook.xml.rels
            this.excel.file("xl/_rels/workbook.xml.rels", builder.buildObject(this.workbookxml_rels));
            //workbook.xml
            this.excel.file("xl/workbook.xml", builder.buildObject(this.workbookxml));
            //sheet_xmls
            _underscore2['default'].each(this.sheet_xmls, function (sheet) {
                if (sheet.name) {
                    var sheet_obj = {};
                    sheet_obj.worksheet = {};
                    _underscore2['default'].extend(sheet_obj.worksheet, sheet.worksheet);
                    _this5.excel.file('xl/worksheets/' + sheet.name, builder.buildObject(sheet_obj));
                }
            });
            //call JSZip#generate()
            return this.excel.generate(option);
        }

        /**
         * * _simple_render
         * * @param {Object} bind_data binding data
         * * @param {Object} jszip_option JsZip#generate() option.
         * * @returns {Object} rendered MS-Excel data. data-format is determined by jszip_option
         * * @private
         **/
    }, {
        key: '_simple_render',
        value: function _simple_render(bind_data, jszip_option) {
            return this.excel.file('xl/sharedStrings.xml', _mustache2['default'].render(this.sharedstrings_str, bind_data)).generate(jszip_option);
        }

        /**
         * * _parse_common_string_with_variable
         * * @return {Array} including common strings only having mustache-variable
         * * @private
         **/
    }, {
        key: '_parse_common_string_with_variable',
        value: function _parse_common_string_with_variable() {
            var _this6 = this;

            var common_strings_with_variable = [];

            _underscore2['default'].each(this.sharedstrings, function (string_obj, index) {
                if ((0, _underscore2['default'])((0, _underscore2['default'])(string_obj.t).string_value()).has_variable()) {
                    string_obj.shared_index = index;
                    common_strings_with_variable.push(string_obj);
                }
            });
            _underscore2['default'].each(common_strings_with_variable, function (common_string_with_variable) {
                common_string_with_variable.using_cells = [];
                _underscore2['default'].each(_this6.template_sheet_data, function (row) {
                    _underscore2['default'].each(row.c, function (cell) {
                        if (cell['$'].t === 's') {
                            if (common_string_with_variable.shared_index === parseInt(cell.v[0])) {
                                common_string_with_variable.using_cells.push(cell['$'].r);
                            }
                        }
                    });
                });
            });

            return common_strings_with_variable;
        }

        /**
         * * _parse_dir_in_excel
         * * @param {String} dir directory name in Zip file.
         * * @return {Promise|Array} array including files parsed by xml2js
         * * @private
         **/
    }, {
        key: '_parse_dir_in_excel',
        value: function _parse_dir_in_excel(dir) {
            var _this7 = this;

            var files = this.excel.folder(dir).file(/.xml/);
            var file_xmls = [];
            return files.reduce(function (promise, file) {
                return promise.then(function (prior_file) {
                    return _bluebird2['default'].resolve().then(function () {
                        return parseString(_this7.excel.file(file.name).asText());
                    }).then(function (file_xml) {
                        file_xml.name = file.name.split('/')[file.name.split('/').length - 1];
                        file_xmls.push(file_xml);
                        return file_xmls;
                    });
                });
            }, _bluebird2['default'].resolve());
        }

        /**
         * * _build_new_sheet
         * * @param {Object} source_sheet
         * * @param {Array} common_strings_with_variable
         * * @return {Object}
         * * @private
         **/
    }, {
        key: '_build_new_sheet',
        value: function _build_new_sheet(source_sheet, common_strings_with_variable) {
            var added_sheet = JSON.parse(JSON.stringify(source_sheet));
            _underscore2['default'].each(common_strings_with_variable, function (e, index) {
                _underscore2['default'].each(e.using_cells, function (cell_address) {
                    _underscore2['default'].each(added_sheet.worksheet.sheetData[0].row, function (row) {
                        _underscore2['default'].each(row.c, function (cell) {
                            if (cell['$'].r === cell_address) {
                                cell.v[0] = e.shared_index;
                            }
                        });
                    });
                });
            });
            return added_sheet;
        }

        /**
         * * _available_sheetid
         * * @return {String} id of next sheet
         * * @private
         **/
    }, {
        key: '_available_sheetid',
        value: function _available_sheetid() {
            var max_rel = _underscore2['default'].max(this.workbookxml_rels.Relationships.Relationship, function (e) {
                return Number(e['$'].Id.replace('rId', ''));
            });
            var next_id = 'rId' + ('00' + (parseInt(max_rel['$'].Id.replace('rId', '')) + parseInt(1))).slice(-3);
            return next_id;
        }

        /**
         * * _sheet_by_name
         * * @param {String} sheetname target sheet name
         * * @return {Object} sheet object
         * * @private
         **/
    }, {
        key: '_sheet_by_name',
        value: function _sheet_by_name(sheetname) {
            var sheetid = _underscore2['default'].find(this.workbookxml.workbook.sheets[0].sheet, function (e) {
                return e['$'].name === sheetname;
            })['$']['r:id'];
            var target_file_path = _underscore2['default'].max(this.workbookxml_rels.Relationships.Relationship, function (e) {
                return e['$'].Id === sheetid;
            })['$'].Target;
            var target_file_name = target_file_path.split('/')[target_file_path.split('/').length - 1];
            var sheet_xml = _underscore2['default'].find(this.sheet_xmls, function (e) {
                return e.name === target_file_name;
            });
            var sheet = { path: target_file_path, value: sheet_xml };
            return sheet;
        }
    }]);

    return SpreadSheet;
})();

module.exports = SpreadSheet;