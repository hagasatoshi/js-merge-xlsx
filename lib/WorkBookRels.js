/**
 * WorkBookRels.js
 * Manage 'xl/_rels/workbook.xml.rels'
 * @author Satoshi Haga
 * @date 2016/04/03
 */

'use strict';

var _createClass = (function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ('value' in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; })();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError('Cannot call a class as a function'); } }

var _ = require('underscore');
var config = require('./Config');

var WorkBookRels = (function () {
    function WorkBookRels(workBookRels) {
        _classCallCheck(this, WorkBookRels);

        this.workBookRels = workBookRels;
        this.sheetRelationships = workBookRels.Relationships.Relationship;
    }

    _createClass(WorkBookRels, [{
        key: 'value',
        value: function value() {
            this.workBookRels.Relationships.Relationship = this.sheetRelationships;
            return this.workBookRels;
        }
    }, {
        key: 'add',
        value: function add(sheetId) {
            this.sheetRelationships.push({
                '$': {
                    Id: sheetId,
                    Type: config.OPEN_XML_SCHEMA_DEFINITION,
                    Target: 'worksheets/sheet' + sheetId + '.xml'
                }
            });
            return this;
        }
    }, {
        key: 'delete',
        value: function _delete(sheetPath) {
            var _this = this;

            _.each(this.sheetRelationships, function (sheet, index) {
                if (sheet && sheet['$'].Target === sheetPath) {
                    _this.sheetRelationships.splice(index, 1);
                }
            });
        }
    }, {
        key: 'findSheetPath',
        value: function findSheetPath(sheetId) {
            var found = _.find(this.sheetRelationships, function (e) {
                return e['$'].Id === sheetId;
            });
            return found ? found['$'].Target : null;
        }
    }, {
        key: 'nextRelationshipId',
        value: function nextRelationshipId() {
            var maxRel = _.max(this.sheetRelationships, function (e) {
                return Number(e['$'].Id.replace('rId', ''));
            });
            var nextId = 'rId' + ('00' + ((maxRel['$'].Id.replace('rId', '') >> 0) + 1)).slice(-3);
            return nextId;
        }
    }]);

    return WorkBookRels;
})();

module.exports = WorkBookRels;