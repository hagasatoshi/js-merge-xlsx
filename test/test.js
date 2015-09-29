/**
 * * test.js
 * * Test script for js-merge-xlsx
 * * @author Satoshi Haga
 * * @date 2015/09/30
 **/

'use strict';

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { 'default': obj }; }

var _assert = require('assert');

var _assert2 = _interopRequireDefault(_assert);

var _excelmerge = require('../excelmerge');

var _excelmerge2 = _interopRequireDefault(_excelmerge);

describe('sampleTest', function () {
    describe('sample', function () {
        it('this is sample src_test', function () {
            var merge = new _excelmerge2['default']('1');
            _assert2['default'].equal(merge.excel, '1');
        });
    });
});