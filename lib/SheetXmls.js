/**
 * SheetXmls.js
 * Manage 'xl/worksheets/*.xml'
 * @author Satoshi Haga
 * @date 2016/04/03
 */

'use strict';

var _createClass = (function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ('value' in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; })();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError('Cannot call a class as a function'); } }

var _ = require('underscore');
var SheetModel = require('./SheetModel');

var SheetXmls = (function () {
    function SheetXmls(sheetObjs) {
        _classCallCheck(this, SheetXmls);

        this.sheetModels = _.map(sheetObjs, function (sheetObj) {
            return new SheetModel(sheetObj);
        });
    }

    _createClass(SheetXmls, [{
        key: 'value',
        value: function value() {
            return _.map(this.sheetModels, function (sheetModel) {
                if (!sheetModel.getName()) {
                    return null;
                }
                var wrk = {};
                wrk.worksheet = {};
                _.extend(wrk.worksheet, sheetModel.value().worksheet);
                return { name: sheetModel.getName(), data: wrk };
            });
        }
    }, {
        key: 'names',
        value: function names() {
            return _.map(this.sheetModels, function (sheetModel) {
                return sheetModel.getName();
            });
        }
    }, {
        key: 'add',
        value: function add(xmlFileName, sheetModel) {
            sheetModel.setName(xmlFileName);
            this.sheetModels.push(sheetModel);
        }
    }, {
        key: 'delete',
        value: function _delete(fileName) {
            _.splice(this.sheetModels, function (sheetModel) {
                return sheetModel.value() && sheetModel.getName() === fileName;
            });
        }
    }, {
        key: 'find',
        value: function find(fileName) {
            return _.find(this.sheetModels, function (sheetModel) {
                return sheetModel.value() && sheetModel.getName() === fileName;
            });
        }
    }, {
        key: 'getTemplateSheetModel',
        value: function getTemplateSheetModel() {
            return _.find(this.sheetModels, function (sheetModel) {
                return sheetModel.isValid();
            });
        }
    }, {
        key: 'stringCount',
        value: function stringCount() {
            return _.sum(this.sheetModels, function (sheetModel) {
                return sheetModel.stringCount();
            });
        }
    }, {
        key: 'templateSheetData',
        value: function templateSheetData() {
            return _.find(this.sheetModels, function (sheetModel) {
                return sheetModel.getName().indexOf('.rels') === -1;
            }).value().worksheet.sheetData[0].row;
        }
    }]);

    return SheetXmls;
})();

module.exports = SheetXmls;