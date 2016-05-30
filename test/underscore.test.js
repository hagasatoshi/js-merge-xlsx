'use strict';

var _ = require('underscore');
require('../lib/underscore');
var assert = require('chai').assert;

describe('underscore.js', function () {
    describe('consistOf()', function () {

        it('should return true if element have all properies', function () {
            var testObj = { field1: 'value1', field2: 'value2', field3: 'value3' };
            assert.isOk(_.consistOf(testObj, ['field1', 'field2', 'field3']));
        });
    });
});