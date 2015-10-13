/**
 * * SpreadSheet
 * * Manage MS-Excel file. core business-logic class for js-merge-xlsx.
 * * @author Satoshi Haga
 * * @date 2015/10/03
 **/
//FIXME count in sharedstring.xml is not correct. fix count.
'use strict';

var _createClass = (function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ('value' in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; })();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError('Cannot call a class as a function'); } }

var Mustache = require('mustache');
var Promise = require('bluebird');
var _ = require('underscore');
require('./underscore_mixin');
var JSZip = require('jszip');
var isNode = require('detect-node');
var outputBuffer = { type: isNode ? 'nodebuffer' : 'blob', compression: "DEFLATE" };
var jszipBuffer = { type: isNode ? 'nodebuffer' : 'arraybuffer', compression: "DEFLATE" };
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
         * * commonStringsWithVariable {Array} including common strings only having mustache variables
         * * sheetXmls {Array} including objects parsed from  'xl/worksheets/*.xml'
         * * sheetXmlsRels {Array} including objects pared from 'xl/worksheets/_rels/*.xml.rels'
         * * templateSheetData {Object} object parsed from 'xl/worksheets/*.xml'. this is used as template-file
         * * templateSheetName {String} sheet-name of template-file
         * * workbookxmlRels {Object} parsed from 'xl/_rels/workbook.xml.rels'
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
            this.commonStringsWithVariable = [];

            //some members are parsed in promise-chain because xml2js parses asynchronously
            return Promise.props({
                sharedstringsObj: parseString(excel.file('xl/sharedStrings.xml').asText()),
                workbookxmlRels: parseString(this.excel.file('xl/_rels/workbook.xml.rels').asText()),
                workbookxml: parseString(this.excel.file('xl/workbook.xml').asText()),
                sheetXmls: this._parseDirInExcel('xl/worksheets'),
                sheetXmlsRels: this._parseDirInExcel('xl/worksheets/_rels')
            }).then(function (_ref) {
                var sharedstringsObj = _ref.sharedstringsObj;
                var workbookxmlRels = _ref.workbookxmlRels;
                var workbookxml = _ref.workbookxml;
                var sheetXmls = _ref.sheetXmls;
                var sheetXmlsRels = _ref.sheetXmlsRels;

                _this.sharedstrings = sharedstringsObj.sst.si;
                _this.workbookxmlRels = workbookxmlRels;
                _this.workbookxml = workbookxml;
                _this.sheetXmls = sheetXmls;
                _this.sheetXmlsRels = sheetXmlsRels;
                _this.templateSheetName = _this.workbookxml.workbook.sheets[0].sheet[0]['$'].name;
                _this.templateSheetData = _.find(sheetXmls, function (e) {
                    return e.name.indexOf('.rels') === -1;
                }).worksheet.sheetData[0].row;
                _this.templateSheetRelsData = _(_this._templateSheetRels()).deepCopy();
                _this.commonStringsWithVariable = _this._parseCommonStringWithVariable();
                //return this for chaining
                return _this;
            });
        }

        /**
         * * simpleRender
         * * @param {Object} bind_data binding data
         * * @returns {Promise|Object} rendered MS-Excel data. data-format is determined by jszip_option
         **/
    }, {
        key: 'simpleRender',
        value: function simpleRender(bindData) {
            var _this2 = this;

            //validation
            if (!bindData) return Promise.reject('simpleRender() must has parameter');

            return Promise.resolve().then(function () {
                return _this2._simpleRender(bindData, outputBuffer);
            });
        }

        /**
         * * bulkRenderMultiFile
         * * @param {Array} bindDataArray including data{name: file's name, data: binding-data}
         * * @returns {Promise|Object} rendered MS-Excel data.
         **/
    }, {
        key: 'bulkRenderMultiFile',
        value: function bulkRenderMultiFile(bindDataArray) {
            var _this3 = this;

            //validation
            if (!_.isArray(bindDataArray)) return Promise.reject('bulkRenderMultiFile() has only array object');
            if (_.find(bindDataArray, function (e) {
                return !(e.name && e.data);
            })) return Promise.reject('bulkRenderMultiFile() is called with invalid parameter');

            var allExcels = new JSZip();
            _.each(bindDataArray, function (_ref2) {
                var name = _ref2.name;
                var data = _ref2.data;
                return allExcels.file(name, _this3._simpleRender(data, jszipBuffer));
            });
            return Promise.resolve().then(function () {
                return allExcels.generate(outputBuffer);
            });
        }

        /**
         * * addSheetBindingData
         * * @param {String} dest_sheet_name name of new sheet
         * * @param {Object} data binding data
         * * @return {Object} this instance for chaining
         **/
    }, {
        key: 'addSheetBindingData',
        value: function addSheetBindingData(destSheetName, data) {
            var _this4 = this;

            //validation
            if (!destSheetName || !data) return Promise.reject('addSheetBindingData() needs to have 2 paramter.');
            //1.add relation of next sheet
            var nextId = this._availableSheetid();
            this.workbookxmlRels.Relationships.Relationship.push({ '$': { Id: nextId, Type: OPEN_XML_SCHEMA_DEFINITION, Target: 'worksheets/sheet' + nextId + '.xml' } });
            this.workbookxml.workbook.sheets[0].sheet.push({ '$': { name: destSheetName, sheetId: nextId.replace('rId', ''), 'r:id': nextId } });

            //2.add sheet file.
            //2-1.prepare rendered-strings
            var renderedStrings = _(this.commonStringsWithVariable).deepCopy();
            _.each(renderedStrings, function (e) {
                return e.t[0] = Mustache.render(_(e.t).stringValue(), data);
            });

            //2-2.add rendered-string into sharedstrings
            var currentCount = this.sharedstrings.length;
            _.each(renderedStrings, function (e, index) {
                e.sharedIndex = currentCount + index;
                _this4.sharedstrings.push(e);
            });

            //2-4.build new sheet oject
            var sourceSheet = this._sheetByName(this.templateSheetName).value;
            var addedSheet = this._buildNewSheet(sourceSheet, renderedStrings);

            //2-5.update sheet name.
            addedSheet.name = 'sheet' + nextId + '.xml';

            //2-6.add this sheet into sheet_xmls
            this.sheetXmls.push(addedSheet);

            return this;
        }

        /**
         * * activateSheet
         * * @param {String} sheetname target sheet name
         * * @return {Object} this instance for chaining
         **/
    }, {
        key: 'activateSheet',
        value: function activateSheet(sheetname) {

            //validation
            if (!sheetname) return Promise.reject('activateSheet() needs to have 1 paramter.');

            var targetSheetName = this._sheetByName(sheetname);
            if (!targetSheetName) return Promise.reject('Invalid sheet name \'' + sheetname + '\'.');

            _.each(this.sheet_xmls, function (sheet) {
                if (!sheet.worksheet) return;
                sheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected = sheet.name === targetSheetName.value.worksheet.name ? '1' : '0';
            });
            return this;
        }

        /**
         * * forcusOnFirstSheet
         * * @return {Object} this instance for chaining
         **/
    }, {
        key: 'forcusOnFirstSheet',
        value: function forcusOnFirstSheet() {
            return this.activateSheet(this._firstSheetName());
        }

        /**
         * * deleteSheet
         * * @param {String} sheetname target sheet name
         * * @return {Object} this instance for chaining
         **/
    }, {
        key: 'deleteSheet',
        value: function deleteSheet(sheetname) {
            var _this5 = this;

            if (!sheetname) return Promise.reject('deleteSheet() needs to have 1 paramter.');
            var targetSheet = this._sheetByName(sheetname);
            if (!targetSheet) return Promise.reject('Invalid sheet name \'' + sheetname + '\'.');
            _.each(this.workbookxmlRels.Relationships.Relationship, function (sheet, index) {
                if (sheet && sheet['$'].Target === targetSheet.path) _this5.workbookxmlRels.Relationships.Relationship.splice(index, 1);
            });
            _.each(this.workbookxml.workbook.sheets[0].sheet, function (sheet, index) {
                if (sheet && sheet['$'].name === sheetname) _this5.workbookxml.workbook.sheets[0].sheet.splice(index, 1);
            });
            _.each(this.sheetXmls, function (sheetXml, index) {
                if (sheetXml && sheetXml.name === targetSheet.value.name) {
                    _this5.sheetXmls.splice(index, 1);
                    _this5.excel.remove('xl/worksheets/' + targetSheet.value.name);
                    _this5.excel.remove('xl/worksheets/_rels/' + targetSheet.value.name + '.rels');
                }
            });
            return this;
        }

        /**
         * * deleteTemplateSheet
         * * @return {Object} this instance for chaining
         **/
    }, {
        key: 'deleteTemplateSheet',
        value: function deleteTemplateSheet() {
            return this.deleteSheet(this.templateSheetName);
        }

        /**
         * * hasAsSharedString
         * * @param {String} targetStr
         * * @return {boolean}
         **/
    }, {
        key: 'hasAsSharedString',
        value: function hasAsSharedString(targetStr) {
            return this.excel.file('xl/sharedStrings.xml').asText().indexOf(targetStr) !== -1;
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

            return parseString(this.excel.file('xl/sharedStrings.xml').asText()).then(function (sharedstringsObj) {

                //sharedstring
                sharedstringsObj.sst.si = _this6._cleanSharedStrings();
                sharedstringsObj.sst['$'].count = sharedstringsObj.sst['$'].uniqueCount = _this6.sharedstrings.length;
                _this6.excel.file('xl/sharedStrings.xml', builder.buildObject(sharedstringsObj)).file("xl/_rels/workbook.xml.rels", builder.buildObject(_this6.workbookxmlRels)).file("xl/workbook.xml", builder.buildObject(_this6.workbookxml));

                //sheetXmls
                _.each(_this6.sheetXmls, function (sheet) {
                    if (sheet.name) {
                        var sheetObj = {};
                        sheetObj.worksheet = {};
                        _.extend(sheetObj.worksheet, sheet.worksheet);
                        _this6.excel.file('xl/worksheets/' + sheet.name, builder.buildObject(sheetObj));
                    }
                });

                //sheetXmlsRels
                if (_this6.templateSheetRelsData.value && _this6.templateSheetRelsData.value.Relationships) {
                    (function () {
                        var strTemplateSheetRels = builder.buildObject({ Relationships: _this6.templateSheetRelsData.value.Relationships });
                        _.each(_this6.sheetXmls, function (sheet) {
                            if (sheet.name) _this6.excel.file('xl/worksheets/_rels/' + sheet.name + '.rels', strTemplateSheetRels);
                        });
                    })();
                }

                //call JSZip#generate()
                return _this6.excel.generate(option);
            });
        }

        /**
         * * _simpleRender
         * * @param {Object} bindData binding data
         * * @param {Object} option JsZip#generate() option.
         * * @returns {Object} rendered MS-Excel data. data-format is determined by jszip_option
         * * @private
         **/
    }, {
        key: '_simpleRender',
        value: function _simpleRender(bindData) {
            var option = arguments.length <= 1 || arguments[1] === undefined ? outputBuffer : arguments[1];

            return new JSZip(this.excel.generate(jszipBuffer)).file('xl/sharedStrings.xml', Mustache.render(this.excel.file('xl/sharedStrings.xml').asText(), bindData)).generate(option);
        }

        /**
         * * _parseCommonStringWithVariable
         * * @return {Array} including common strings only having mustache-variable
         * * @private
         **/
    }, {
        key: '_parseCommonStringWithVariable',
        value: function _parseCommonStringWithVariable() {
            var _this7 = this;

            var commonStringsWithVariable = [];
            _.each(this.sharedstrings, function (stringObj, index) {
                if (_(stringObj.t).stringValue() && _(_(stringObj.t).stringValue()).hasVariable()) {
                    stringObj.sharedIndex = index;
                    commonStringsWithVariable.push(stringObj);
                }
            });
            _.each(commonStringsWithVariable, function (commonStringWithVariable) {
                commonStringWithVariable.usingCells = [];
                _.each(_this7.templateSheetData, function (row) {
                    _.each(row.c, function (cell) {
                        if (cell['$'].t === 's') {
                            if (commonStringWithVariable.sharedIndex === cell.v[0] >> 0) {
                                commonStringWithVariable.usingCells.push(cell['$'].r);
                            }
                        }
                    });
                });
            });

            return commonStringsWithVariable;
        }

        /**
         * * _parseDirInExcel
         * * @param {String} dir directory name in Zip file.
         * * @return {Promise|Array} array including files parsed by xml2js
         * * @private
         **/
    }, {
        key: '_parseDirInExcel',
        value: function _parseDirInExcel(dir) {
            var _this8 = this;

            var files = this.excel.folder(dir).file(/.xml/);
            var fileXmls = [];
            return files.reduce(function (promise, file) {
                return promise.then(function (prior_file) {
                    return Promise.resolve().then(function () {
                        return parseString(_this8.excel.file(file.name).asText());
                    }).then(function (file_xml) {
                        file_xml.name = _.last(file.name.split('/'));
                        fileXmls.push(file_xml);
                        return fileXmls;
                    });
                });
            }, Promise.resolve());
        }

        /**
         * * _buildNewSheet
         * * @param {Object} sourceSheet
         * * @param {Array} commonStringsWithVariable
         * * @return {Object}
         * * @private
         **/
    }, {
        key: '_buildNewSheet',
        value: function _buildNewSheet(sourceSheet, commonStringsWithVariable) {
            var addedSheet = _(sourceSheet).deepCopy();
            _.each(commonStringsWithVariable, function (e, index) {
                _.each(e.usingCells, function (cellAddress) {
                    _.each(addedSheet.worksheet.sheetData[0].row, function (row) {
                        _.each(row.c, function (cell) {
                            if (cell['$'].r === cellAddress) {
                                cell.v[0] = e.sharedIndex;
                            }
                        });
                    });
                });
            });
            addedSheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected = '0';
            return addedSheet;
        }

        /**
         * * _availableSheetid
         * * @return {String} id of next sheet
         * * @private
         **/
    }, {
        key: '_availableSheetid',
        value: function _availableSheetid() {
            var maxRel = _.max(this.workbookxmlRels.Relationships.Relationship, function (e) {
                return Number(e['$'].Id.replace('rId', ''));
            });
            var nextId = 'rId' + ('00' + ((maxRel['$'].Id.replace('rId', '') >> 0) + 1)).slice(-3);
            return nextId;
        }

        /**
         * * _sheetByName
         * * @param {String} sheetname target sheet name
         * * @return {Object} sheet object
         * * @private
         **/
    }, {
        key: '_sheetByName',
        value: function _sheetByName(sheetname) {
            var targetSheet = _.find(this.workbookxml.workbook.sheets[0].sheet, function (e) {
                return e['$'].name === sheetname;
            });
            if (!targetSheet) return null; //invalid sheet name

            var sheetid = targetSheet['$']['r:id'];
            var targetFilePath = _.max(this.workbookxmlRels.Relationships.Relationship, function (e) {
                return e['$'].Id === sheetid;
            })['$'].Target;
            var targetFileName = _.last(targetFilePath.split('/'));
            return { path: targetFilePath, value: _.find(this.sheetXmls, function (e) {
                    return e.name === targetFileName;
                }) };
        }

        /**
         * * _sheetRelsByName
         * * @param {String} sheetname target sheet name
         * * @return {Object} sheet_rels object
         * * @private
         **/
    }, {
        key: '_sheetRelsByName',
        value: function _sheetRelsByName(sheetname) {
            var targetFilePath = this._sheetByName(sheetname).path;
            var targetName = _.last(targetFilePath.split('/')) + '.rels';
            return { name: targetName, value: _.find(this.sheetXmlsRels, function (e) {
                    return e.name === targetName;
                }) };
        }

        /**
         * * _templateSheetRels
         * * @return {Object} sheet_rels object of template-sheet
         * * @private
         **/
    }, {
        key: '_templateSheetRels',
        value: function _templateSheetRels() {
            return this._sheetRelsByName(this.templateSheetName);
        }

        /**
         * * _sheetNames
         * * @return {Array} array including sheet name
         * * @private
         **/
    }, {
        key: '_sheetNames',
        value: function _sheetNames() {
            return _.map(this.sheetXmls, function (e) {
                return e.name;
            });
        }

        /**
         * * _cleanSharedStrings
         * * @return {Array} shared strings
         * * @private
         **/
    }, {
        key: '_cleanSharedStrings',
        value: function _cleanSharedStrings() {
            _.each(this.sharedstrings, function (e) {
                delete e.sharedIndex;
                delete e.usingCells;
            });
            return this.sharedstrings;
        }

        /**
         * * _firstSheetName
         * * @return {String} name of first-sheet of MS-Excel file
         * * @private
         **/
    }, {
        key: '_firstSheetName',
        value: function _firstSheetName() {
            return this.workbookxml.workbook.sheets[0].sheet[0]['$'].name;
        }

        /**
         * * _activeSheets
         * * @return {Array} array including only active sheets.
         * * @private
         **/
    }, {
        key: '_activeSheets',
        value: function _activeSheets() {
            return _.filter(this.sheetXmls, function (sheet) {
                return sheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected === '1';
            });
        }

        /**
         * * _deactiveSheets
         * * @return {Array} array including only deactive sheets.
         * * @private
         **/
    }, {
        key: '_deactiveSheets',
        value: function _deactiveSheets() {
            return _.filter(this.sheetXmls, function (sheet) {
                return sheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected === '0';
            });
        }
    }]);

    return SpreadSheet;
})();

module.exports = SpreadSheet;