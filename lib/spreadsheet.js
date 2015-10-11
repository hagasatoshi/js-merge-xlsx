/**
 * * SpreadSheet
 * * Manage MS-Excel file. core business-logic class for js-merge-xlsx.
 * * @author Satoshi Haga
 * * @date 2015/10/03
 **/
'use strict';

var _createClass = (function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ('value' in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; })();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError('Cannot call a class as a function'); } }

var Mustache = require('mustache');
var Promise = require('bluebird');
var _ = require('underscore');
require('./underscore_mixin');
var JSZip = require('jszip');
var isNode = require('detect-node');
var output_buffer = { type: isNode ? 'nodebuffer' : 'blob', compression: "DEFLATE" };
var jszip_buffer = { type: isNode ? 'nodebuffer' : 'arraybuffer', compression: "DEFLATE" };
var xml2js = require('xml2js');
var parseString = Promise.promisify(xml2js.parseString);
var builder = new xml2js.Builder();

var OPEN_XML_SCHEMA_DEFINITION = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';

var SpreadSheet = (function () {
    function SpreadSheet() {
        _classCallCheck(this, SpreadSheet);
    }

    //Exports

    _createClass(SpreadSheet, [{
        key: 'load',

        /**
         * * member variables
         * * excel {Object} JSZip instance including template excel file
         * * variables {Array} including mustache-variables defined in sharedstrings.xml
         * * sharedstrings {Array} includings common strings defined in sharedstrings.xml
         * * sharedstrings_obj {Object} whole sharedstrings object
         * * common_strings_with_variable {Array} including common strings only having mustache variables
         * * sheet_xmls {Array} including objects parsed from  'xl/worksheets/*.xml'
         * * sheet_xmls_rels {Array} including objects pared from 'xl/worksheets/_rels/*.xml.rels'
         * * template_sheet_data {Object} object parsed from 'xl/worksheets/*.xml'. this is used as template-file
         * * template_sheet_name {String} sheet-name of template-file
         * * workbookxml_rels {Object} parsed from 'xl/_rels/workbook.xml.rels'
         * * workbookxml {Object} parsed from 'xl/workbook.xml'
         * */

        /**
         * * load
         * * @param {Object} excel JsZip object including MS-Excel file
         * * @return {Promise|Object} Promise instance including this
         **/
        value: function load(excel) {
            var _this = this;

            //validation
            if (!(excel instanceof JSZip)) return Promise.reject('First parameter must be JSZip instance including MS-Excel data');

            //set member variable
            this.excel = excel;
            this.variables = _(excel.file('xl/sharedStrings.xml').asText()).variables();
            this.common_strings_with_variable = [];

            //some members are parsed in promise-chain because xml2js parses asynchronously
            return Promise.props({
                sharedstrings_obj: parseString(excel.file('xl/sharedStrings.xml').asText()),
                workbookxml_rels: parseString(this.excel.file('xl/_rels/workbook.xml.rels').asText()),
                workbookxml: parseString(this.excel.file('xl/workbook.xml').asText()),
                sheet_xmls: this._parse_dir_in_excel('xl/worksheets'),
                sheet_xmls_rels: this._parse_dir_in_excel('xl/worksheets/_rels')
            }).then(function (templates) {
                _this.sharedstrings = templates.sharedstrings_obj.sst.si;
                _this.workbookxml_rels = templates.workbookxml_rels;
                _this.workbookxml = templates.workbookxml;
                _this.sheet_xmls = templates.sheet_xmls;
                _this.sheet_xmls_rels = templates.sheet_xmls_rels;
                _this.template_sheet_name = _this.workbookxml.workbook.sheets[0].sheet[0]['$'].name;
                _this.template_sheet_data = _.find(templates.sheet_xmls, function (e) {
                    return e.name.indexOf('.rels') === -1;
                }).worksheet.sheetData[0].row;
                _this.template_sheet_rels_data = _(_this._template_sheet_rels()).deep_copy();
                _this.common_strings_with_variable = _this._parse_common_string_with_variable();

                //return this for chaining
                return _this;
            });
        }

        /**
         * * simple_render
         * * @param {Object} bind_data binding data
         * * @returns {Promise|Object} rendered MS-Excel data. data-format is determined by jszip_option
         **/
    }, {
        key: 'simple_render',
        value: function simple_render(bind_data) {
            var _this2 = this;

            //validation
            if (!bind_data) return Promise.reject('simple_render() must has parameter');

            return Promise.resolve().then(function () {
                return _this2._simple_render(bind_data, output_buffer);
            });
        }

        /**
         * * bulk_render_multi_file
         * * @param {Array} bind_data_array including data{name: file's name, data: binding-data}
         * * @returns {Promise|Object} rendered MS-Excel data.
         **/
    }, {
        key: 'bulk_render_multi_file',
        value: function bulk_render_multi_file(bind_data_array) {
            var _this3 = this;

            //validation
            if (!_.isArray(bind_data_array)) return Promise.reject('bulk_render_multi_file() has only array object');
            if (_.find(bind_data_array, function (e) {
                return !(e.name && e.data);
            })) return Promise.reject('bulk_render_multi_file() is called with invalid parameter');

            var all_excels = new JSZip();
            _.each(bind_data_array, function (_ref) {
                var name = _ref.name;
                var data = _ref.data;
                return all_excels.file(name, _this3._simple_render(data, jszip_buffer));
            });
            return Promise.resolve().then(function () {
                return all_excels.generate(output_buffer);
            });
        }

        /**
         * * add_sheet_binding_data
         * * @param {String} dest_sheet_name name of new sheet
         * * @param {Object} data binding data
         * * @return {Object} this instance for chaining
         **/
    }, {
        key: 'add_sheet_binding_data',
        value: function add_sheet_binding_data(dest_sheet_name, data) {
            var _this4 = this;

            //validation
            if (!dest_sheet_name || !data) return Promise.reject('add_sheet_binding_data() needs to have 2 paramter.');

            //1.add relation of next sheet
            var next_id = this._available_sheetid();
            this.workbookxml_rels.Relationships.Relationship.push({ '$': { Id: next_id, Type: OPEN_XML_SCHEMA_DEFINITION, Target: 'worksheets/sheet' + next_id + '.xml' } });
            this.workbookxml.workbook.sheets[0].sheet.push({ '$': { name: dest_sheet_name, sheetId: next_id.replace('rId', ''), 'r:id': next_id } });

            //2.add sheet file.
            //2-1.prepare rendered-strings
            var rendered_strings = _(this.common_strings_with_variable).deep_copy();
            _.each(rendered_strings, function (e) {
                return e.t[0] = Mustache.render(_(e.t).string_value(), data);
            });

            //2-2.add rendered-string into sharedstrings
            var current_count = this.sharedstrings.length;
            _.each(rendered_strings, function (e, index) {
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

            return this;
        }

        /**
         * * activate_sheet
         * * @param {String} sheetname target sheet name
         * * @return {Object} this instance for chaining
         **/
    }, {
        key: 'activate_sheet',
        value: function activate_sheet(sheetname) {

            //validation
            if (!sheetname) return Promise.reject('activate_sheet() needs to have 1 paramter.');

            var target_sheet_name = this._sheet_by_name(sheetname);
            if (!target_sheet_name) return Promise.reject('Invalid sheet name \'' + sheetname + '\'.');

            _.each(this.sheet_xmls, function (sheet) {
                if (!sheet.worksheet) return;
                sheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected = sheet.name === target_sheet_name.value.worksheet.name ? '1' : '0';
            });
            return this;
        }

        /**
         * * forcus_on_first_sheet
         * * @return {Object} this instance for chaining
         **/
    }, {
        key: 'forcus_on_first_sheet',
        value: function forcus_on_first_sheet() {
            return this.activate_sheet(this._first_sheet_name());
        }

        /**
         * * delete_sheet
         * * @param {String} sheetname target sheet name
         * * @return {Object} this instance for chaining
         **/
    }, {
        key: 'delete_sheet',
        value: function delete_sheet(sheetname) {
            var _this5 = this;

            if (!sheetname) return Promise.reject('delete_sheet() needs to have 1 paramter.');

            var target_sheet = this._sheet_by_name(sheetname);
            if (!target_sheet) return Promise.reject('Invalid sheet name \'' + sheetname + '\'.');

            _.each(this.workbookxml_rels.Relationships.Relationship, function (sheet, index) {
                if (sheet && sheet['$'].Target === target_sheet.path) {
                    _this5.workbookxml_rels.Relationships.Relationship.splice(index, 1);
                }
            });
            _.each(this.workbookxml.workbook.sheets[0].sheet, function (sheet, index) {
                if (sheet && sheet['$'].name === sheetname) {
                    _this5.workbookxml.workbook.sheets[0].sheet.splice(index, 1);
                }
            });
            _.each(this.sheet_xmls, function (sheet_xml, index) {
                if (sheet_xml && sheet_xml.name === target_sheet.value.name) {
                    _this5.sheet_xmls.splice(index, 1);
                }
            });
            return this;
        }

        /**
         * * delete_template_sheet
         * * @return {Object} this instance for chaining
         **/
    }, {
        key: 'delete_template_sheet',
        value: function delete_template_sheet() {
            return this.delete_sheet(this.template_sheet_name);
        }

        /**
         * * has_as_shared_string
         * * @param {String} target_str
         * * @return {boolean}
         **/
    }, {
        key: 'has_as_shared_string',
        value: function has_as_shared_string(target_str) {
            return this.excel.file('xl/sharedStrings.xml').asText().indexOf(target_str) !== -1;
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
            var _this6 = this;

            parseString(this.excel.file('xl/sharedStrings.xml').asText()).then(function (sharedstrings_obj) {

                //sharedstring
                sharedstrings_obj.sst.si = _this6._clean_shared_strings();
                sharedstrings_obj.sst['$'].count = sharedstrings_obj.sst['$'].uniqueCount = _this6.sharedstrings.length;
                _this6.excel.file('xl/sharedStrings.xml', builder.buildObject(sharedstrings_obj)).file("xl/_rels/workbook.xml.rels", builder.buildObject(_this6.workbookxml_rels)).file("xl/workbook.xml", builder.buildObject(_this6.workbookxml));

                //sheet_xmls
                _.each(_this6.sheet_xmls, function (sheet) {
                    if (sheet.name) {
                        var sheet_obj = {};
                        sheet_obj.worksheet = {};
                        _.extend(sheet_obj.worksheet, sheet.worksheet);
                        _this6.excel.file('xl/worksheets/' + sheet.name, builder.buildObject(sheet_obj));
                    }
                });

                //sheet_xmls_rels
                var str_template_sheet_rels = builder.buildObject(_this6.template_sheet_rels_data);
                _.each(_this6.sheet_xmls, function (sheet) {
                    if (sheet.name) {
                        _this6.excel.file('xl/worksheets/_rels/' + sheet.name + '.rels', str_template_sheet_rels);
                    }
                });

                //call JSZip#generate()
                return _this6.excel.generate(option);
            });
        }

        /**
         * * _simple_render
         * * @param {Object} bind_data binding data
         * * @param {Object} option JsZip#generate() option.
         * * @returns {Object} rendered MS-Excel data. data-format is determined by jszip_option
         * * @private
         **/
    }, {
        key: '_simple_render',
        value: function _simple_render(bind_data) {
            var option = arguments.length <= 1 || arguments[1] === undefined ? output_buffer : arguments[1];

            return this.excel.file('xl/sharedStrings.xml', Mustache.render(this.excel.file('xl/sharedStrings.xml').asText(), bind_data)).generate(option);
        }

        /**
         * * _parse_common_string_with_variable
         * * @return {Array} including common strings only having mustache-variable
         * * @private
         **/
    }, {
        key: '_parse_common_string_with_variable',
        value: function _parse_common_string_with_variable() {
            var _this7 = this;

            var common_strings_with_variable = [];
            _.each(this.sharedstrings, function (string_obj, index) {
                if (_(string_obj.t).string_value() && _(_(string_obj.t).string_value()).has_variable()) {
                    string_obj.shared_index = index;
                    common_strings_with_variable.push(string_obj);
                }
            });
            _.each(common_strings_with_variable, function (common_string_with_variable) {
                common_string_with_variable.using_cells = [];
                _.each(_this7.template_sheet_data, function (row) {
                    _.each(row.c, function (cell) {
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
            var _this8 = this;

            var files = this.excel.folder(dir).file(/.xml/);
            var file_xmls = [];
            return files.reduce(function (promise, file) {
                return promise.then(function (prior_file) {
                    return Promise.resolve().then(function () {
                        return parseString(_this8.excel.file(file.name).asText());
                    }).then(function (file_xml) {
                        file_xml.name = file.name.split('/')[file.name.split('/').length - 1];
                        file_xmls.push(file_xml);
                        return file_xmls;
                    });
                });
            }, Promise.resolve());
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
            var added_sheet = _(source_sheet).deep_copy();
            _.each(common_strings_with_variable, function (e, index) {
                _.each(e.using_cells, function (cell_address) {
                    _.each(added_sheet.worksheet.sheetData[0].row, function (row) {
                        _.each(row.c, function (cell) {
                            if (cell['$'].r === cell_address) {
                                cell.v[0] = e.shared_index;
                            }
                        });
                    });
                });
            });
            added_sheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected = '0';
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
            var max_rel = _.max(this.workbookxml_rels.Relationships.Relationship, function (e) {
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
            var target_sheet = _.find(this.workbookxml.workbook.sheets[0].sheet, function (e) {
                return e['$'].name === sheetname;
            });
            if (!target_sheet) return null; //invalid sheet name

            var sheetid = target_sheet['$']['r:id'];
            var target_file_path = _.max(this.workbookxml_rels.Relationships.Relationship, function (e) {
                return e['$'].Id === sheetid;
            })['$'].Target;
            var target_file_name = target_file_path.split('/')[target_file_path.split('/').length - 1];
            var sheet_xml = _.find(this.sheet_xmls, function (e) {
                return e.name === target_file_name;
            });
            var sheet = { path: target_file_path, value: sheet_xml };
            return sheet;
        }

        /**
         * * _sheet_rels_by_name
         * * @param {String} sheetname target sheet name
         * * @return {Object} sheet_rels object
         * * @private
         **/
    }, {
        key: '_sheet_rels_by_name',
        value: function _sheet_rels_by_name(sheetname) {
            var target_file_path = this._sheet_by_name(sheetname).path;
            var target_name = target_file_path.split('/')[target_file_path.split('/').length - 1] + '.rels';
            var sheet_xml_rels = _.find(this.sheet_xmls_rels, function (e) {
                return e.name === target_name;
            });
            var sheet = { name: target_name, value: sheet_xml_rels };
            return sheet;
        }

        /**
         * * _template_sheet_rels
         * * @return {Object} sheet_rels object of template-sheet
         * * @private
         **/
    }, {
        key: '_template_sheet_rels',
        value: function _template_sheet_rels() {
            return this._sheet_rels_by_name(this.template_sheet_name);
        }

        /**
         * * _sheet_names
         * * @return {Array} array including sheet name
         * * @private
         **/
    }, {
        key: '_sheet_names',
        value: function _sheet_names() {
            return _.map(this.sheet_xmls, function (e) {
                return e.name;
            });
        }

        /**
         * * _clean_shared_strings
         * * @return {Array} shared strings
         * * @private
         **/
    }, {
        key: '_clean_shared_strings',
        value: function _clean_shared_strings() {
            return _.map(this.sharedstrings, function (e) {
                return { t: e.t, phoneticPr: e.phoneticPr };
            });
        }

        /**
         * * _first_sheet_name
         * * @return {String} name of first-sheet of MS-Excel file
         * * @private
         **/
    }, {
        key: '_first_sheet_name',
        value: function _first_sheet_name() {
            return this.workbookxml.workbook.sheets[0].sheet[0]['$'].name;
        }

        /**
         * * active_sheets
         * * @return {Array} array including only active sheets.
         * * @private
         **/
    }, {
        key: '_active_sheets',
        value: function _active_sheets() {
            return _.filter(this.sheet_xmls, function (sheet) {
                return sheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected === '1';
            });
        }

        /**
         * * deactive_sheets
         * * @return {Array} array including only deactive sheets.
         * * @private
         **/
    }, {
        key: '_deactive_sheets',
        value: function _deactive_sheets() {
            return _.filter(this.sheet_xmls, function (sheet) {
                return sheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected === '0';
            });
        }
    }]);

    return SpreadSheet;
})();

module.exports = SpreadSheet;