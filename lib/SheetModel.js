/**
 * SheetModel.js
 * Manage each 'xl/worksheets/*.xml'
 * @author Satoshi Haga
 * @date 2016/05/06
 */

'use strict';

var _createClass = (function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ('value' in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; })();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError('Cannot call a class as a function'); } }

var _ = require('underscore');

var SheetModel = (function () {
    function SheetModel(sheetObj) {
        _classCallCheck(this, SheetModel);

        this.sheetObj = sheetObj;
    }

    _createClass(SheetModel, [{
        key: 'stringCount',
        value: function stringCount() {
            var stringCountInRow = function stringCountInRow(cells) {
                return _.count(cells, function (cell) {
                    return !!cell['$'].t;
                });
            };
            return this.sheetObj.worksheet ? _.sum(this.getRows(), stringCountInRow) : 0;
        }
    }, {
        key: 'value',
        value: function value() {
            return this.sheetObj;
        }
    }, {
        key: 'setName',
        value: function setName(xmlFileName) {
            this.sheetObj.name = xmlFileName;
            return this;
        }
    }, {
        key: 'isValid',
        value: function isValid() {
            return !!this.sheetObj.worksheet;
        }
    }, {
        key: 'getName',
        value: function getName() {
            return this.sheetObj.name;
        }
    }, {
        key: 'getRows',
        value: function getRows() {
            return this.sheetObj.worksheet.sheetData[0].row;
        }
    }, {
        key: 'clone',
        value: function clone() {
            var sheetSelected = arguments.length <= 0 || arguments[0] === undefined ? false : arguments[0];

            var cloned = new SheetModel(_.deepCopy(this.sheetObj));
            cloned.setSheetSelected(sheetSelected);
            return cloned;
        }
    }, {
        key: 'setSheetSelected',
        value: function setSheetSelected(sheetSelected) {
            this.sheetObj.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected = sheetSelected ? '1' : '0';
        }
    }, {
        key: 'updateStringIndex',
        value: function updateStringIndex(stringModels) {
            var updateIndex = function updateIndex(stringModel, row) {
                _.each(stringModel.usingCells, function (address) {
                    var cellAtThisAddress = _.find(row.c, function (cell) {
                        return cell['$'].r === address;
                    });
                    if (cellAtThisAddress) {
                        cellAtThisAddress.v[0] = stringModel.sharedIndex;
                    }
                });
            };
            _.nestedEach(stringModels, this.getRows(), updateIndex);
            return this;
        }
    }, {
        key: 'cloneWithMergedString',
        value: function cloneWithMergedString(stringModels) {
            return this.clone().updateStringIndex(stringModels);
        }
    }]);

    return SheetModel;
})();

module.exports = SheetModel;