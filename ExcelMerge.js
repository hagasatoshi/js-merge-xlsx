/**
 * ExcelMerge
 * @author Satoshi Haga
 * @date 2016/03/27
 */

'use strict';

var _createClass = (function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ('value' in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; })();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError('Cannot call a class as a function'); } }

var Promise = require('bluebird');
var _ = require('underscore');
var JSZip = require('jszip');
var Mustache = require('mustache');
var Excel = require('./lib/Excel');
var WorkBookXml = require('./lib/WorkBookXml');
var WorkBookRels = require('./lib/WorkBookRels');
var SheetXmls = require('./lib/SheetXmls');
var SharedStrings = require('./lib/SharedStrings');
var Config = require('./lib/Config');
require('./lib/underscore');

var ExcelMerge = {

    merge: function merge(template, data) {
        var oututType = arguments.length <= 2 || arguments[2] === undefined ? Config.JSZIP_OPTION.BUFFER_TYPE_OUTPUT : arguments[2];

        var templateObj = new JSZip(template);
        return templateObj.file(Config.EXCEL_FILES.FILE_SHARED_STRINGS, Mustache.render(templateObj.file(Config.EXCEL_FILES.FILE_SHARED_STRINGS).asText(), data)).generate({ type: oututType, compression: Config.JSZIP_OPTION.COMPLESSION });
    },

    bulkMergeToFiles: function bulkMergeToFiles(template, arrayObj) {
        return _.reduce(arrayObj, function (zip, _ref) {
            var name = _ref.name;
            var data = _ref.data;

            zip.file(name, ExcelMerge.merge(template, data, Config.JSZIP_OPTION.buffer_type_jszip));
            return zip;
        }, new JSZip()).generate({
            type: Config.JSZIP_OPTION.BUFFER_TYPE_OUTPUT,
            compression: Config.JSZIP_OPTION.COMPLESSION
        });
    },

    bulkMergeToSheets: function bulkMergeToSheets(template, arrayObj) {
        return parse(template).then(function (templateObj) {
            var excelObj = new Merge(templateObj).addMergedSheets(arrayObj).deleteTemplateSheet().value();
            return new Excel(template).generateWithData(excelObj);
        });
    }
};

var parse = function parse(template) {
    var templateObj = new Excel(template);
    return Promise.props({
        sharedstrings: templateObj.parseSharedStrings(),
        workbookxmlRels: templateObj.parseWorkbookRels(),
        workbookxml: templateObj.parseWorkbook(),
        sheetXmls: templateObj.parseWorksheetsDir()
    }).then(function (_ref2) {
        var sharedstrings = _ref2.sharedstrings;
        var workbookxmlRels = _ref2.workbookxmlRels;
        var workbookxml = _ref2.workbookxml;
        var sheetXmls = _ref2.sheetXmls;

        var sheetXmlObjs = new SheetXmls(sheetXmls);
        return {
            relationship: new WorkBookRels(workbookxmlRels),
            workbookxml: new WorkBookXml(workbookxml),
            sheetXmls: sheetXmlObjs,
            templateSheetModel: sheetXmlObjs.getTemplateSheetModel(),
            sharedstrings: new SharedStrings(sharedstrings, sheetXmlObjs.templateSheetData())
        };
    });
};

var Merge = (function () {
    function Merge(templateObj) {
        _classCallCheck(this, Merge);

        this.excelObj = templateObj;
    }

    _createClass(Merge, [{
        key: 'addMergedSheets',
        value: function addMergedSheets(dataArray) {
            var _this = this;

            _.each(dataArray, function (_ref3) {
                var name = _ref3.name;
                var data = _ref3.data;
                return _this.addMergedSheet(name, data);
            });
            return this;
        }
    }, {
        key: 'addMergedSheet',
        value: function addMergedSheet(newSheetName, mergeData) {
            var nextId = this.excelObj.relationship.nextRelationshipId();
            this.excelObj.relationship.add(nextId);
            this.excelObj.workbookxml.add(newSheetName, nextId);
            this.excelObj.sheetXmls.add('sheet' + nextId + '.xml', this.excelObj.templateSheetModel.cloneWithMergedString(this.excelObj.sharedstrings.addMergedStrings(mergeData)));
        }
    }, {
        key: 'deleteTemplateSheet',
        value: function deleteTemplateSheet() {
            var sheetname = this.excelObj.workbookxml.firstSheetName();
            var targetSheet = this.findSheetByName(sheetname);
            this.excelObj.relationship['delete'](targetSheet.path);
            this.excelObj.workbookxml['delete'](sheetname);
            return this;
        }
    }, {
        key: 'findSheetByName',
        value: function findSheetByName(sheetname) {
            var sheetid = this.excelObj.workbookxml.findSheetId(sheetname);
            if (!sheetid) {
                return null;
            }
            var targetFilePath = this.excelObj.relationship.findSheetPath(sheetid);
            var targetFileName = _.last(targetFilePath.split('/'));
            return { path: targetFilePath, value: this.excelObj.sheetXmls.find(targetFileName) };
        }
    }, {
        key: 'value',
        value: function value() {
            return this.excelObj;
        }
    }]);

    return Merge;
})();

module.exports = ExcelMerge;