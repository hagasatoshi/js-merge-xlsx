/**
 * WorkBookXml.js
 * Manage 'xl/workbook.xml'
 * @author Satoshi Haga
 * @date 2016/04/02
 */

'use strict';

var _createClass = (function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ('value' in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; })();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError('Cannot call a class as a function'); } }

var _ = require('underscore');

var WorkBookXml = (function () {
    function WorkBookXml(workBookXml) {
        _classCallCheck(this, WorkBookXml);

        this.workBookXml = workBookXml;
        this.sheetDefinitions = workBookXml.workbook.sheets[0].sheet;
    }

    _createClass(WorkBookXml, [{
        key: 'value',
        value: function value() {
            this.workBookXml.workbook.sheets[0].sheet = this.sheetDefinitions;
            return this.workBookXml;
        }
    }, {
        key: 'add',
        value: function add(sheetName, sheetId) {
            this.sheetDefinitions.push({
                '$': {
                    name: sheetName,
                    sheetId: sheetId.replace('rId', ''),
                    'r:id': sheetId
                }
            });
            return this;
        }
    }, {
        key: 'delete',
        value: function _delete(sheetName) {
            var _this = this;

            _.each(this.sheetDefinitions, function (sheet, index) {
                if (sheet && sheet['$'].name === sheetName) {
                    _this.sheetDefinitions.splice(index, 1);
                }
            });
        }
    }, {
        key: 'findSheetId',
        value: function findSheetId(sheetName) {
            var sheet = _.find(this.sheetDefinitions, function (e) {
                return e['$'].name === sheetName;
            });
            return sheet ? sheet['$']['r:id'] : null;
        }
    }, {
        key: 'firstSheetName',
        value: function firstSheetName() {
            return this.sheetDefinitions[0]['$'].name;
        }
    }]);

    return WorkBookXml;
})();

module.exports = WorkBookXml;