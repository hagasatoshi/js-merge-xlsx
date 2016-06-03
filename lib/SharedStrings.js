/**
 * SharedStrings.js
 * Manage 'xl/sharedStrings.xml'
 * @author Satoshi Haga
 * @date 2016/04/03
 */

'use strict';

var _createClass = (function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ('value' in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; })();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError('Cannot call a class as a function'); } }

var _ = require('underscore');
require('./underscore');
var Mustache = require('mustache');

var SharedStrings = (function () {
    function SharedStrings(rawObj, templateSheetObj) {
        _classCallCheck(this, SharedStrings);

        this.rawObj = rawObj;
        this.templateSheetData = templateSheetObj;
        this.stringModels = this.parseStringModels();
    }

    _createClass(SharedStrings, [{
        key: 'parseStringModels',
        value: function parseStringModels() {
            var _this = this;

            if (!this.rawObj.sst || !this.rawObj.sst.si) {
                return null;
            }
            var stringModels = _.deepCopy(this.rawObj.sst.si);
            _.each(this.filterStaticString(stringModels), function (stringModel) {
                stringModel.usingCells = _.reduce(_this.templateSheetData, function (usingCells, row) {
                    return usingCells.concat(SharedStrings.usingCellAddresses(row.c, stringModel.sharedIndex));
                }, []);
            });
            return stringModels;
        }
    }, {
        key: 'add',
        value: function add(newStringModels) {
            var currentCount = this.stringModels.length;
            _.each(newStringModels, function (e, index) {
                e.sharedIndex = currentCount + index;
            });
            this.stringModels = this.stringModels.concat(newStringModels);
            return newStringModels;
        }
    }, {
        key: 'value',
        value: function value() {
            if (!this.stringModels) {
                return null;
            }
            this.rawObj.sst.si = _.deleteProps(this.stringModels, ['sharedIndex', 'usingCells']);
            this.rawObj.sst['$'].uniqueCount = this.stringModels.length;
            this.rawObj.sst['$'].count = this.stringModels.length;
            return this.rawObj;
        }
    }, {
        key: 'filterStaticString',
        value: function filterStaticString(stringModels) {
            var ret = [];
            _.each(stringModels, function (stringObj, index) {
                if (_.stringValue(stringObj.t) && _.hasVariable(_.stringValue(stringObj.t))) {
                    stringObj.sharedIndex = index;
                    ret.push(stringObj);
                }
            });
            return ret;
        }
    }, {
        key: 'hasString',
        value: function hasString() {
            return !!this.stringModels;
        }
    }, {
        key: 'buildNewSharedStrings',
        value: function buildNewSharedStrings(mergedData) {
            return _.reduce(_.deepCopy(this.filterStaticString(this.stringModels)), function (newSharedStrings, templateString) {
                templateString.t[0] = Mustache.render(_.stringValue(templateString.t), mergedData);
                newSharedStrings.push(templateString);
                return newSharedStrings;
            }, []);
        }
    }, {
        key: 'addMergedStrings',
        value: function addMergedStrings(mergedData) {
            if (!this.hasString()) {
                return;
            }
            return this.add(this.buildNewSharedStrings(mergedData));
        }
    }], [{
        key: 'usingCellAddresses',
        value: function usingCellAddresses(cells, stringIndex) {
            return _.reduce(cells, function (addresses, cell) {
                if (cell['$'].t === 's' && stringIndex === cell.v[0] >> 0) {
                    addresses.push(cell['$'].r);
                }
                return addresses;
            }, []);
        }
    }]);

    return SharedStrings;
})();

module.exports = SharedStrings;