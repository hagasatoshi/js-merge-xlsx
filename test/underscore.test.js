'use strict';

var _ = require('underscore');
require('../lib/underscore');
var assert = require('chai').assert;

describe('underscore.js', function () {

    describe('stringValue()', function () {
        it('should return the same value if not array', function () {
            assert.strictEqual(_.stringValue('test'), 'test');
        });
        it('should return first element if array', function () {
            assert.strictEqual(_.stringValue(['first', 'second', 'third']), 'first');
        });
        it('should return attribute "mustache" if have', function () {
            assert.strictEqual(_.stringValue([{ _: 'mustache', key1: 'value1' }]), 'mustache');
        });
    });

    describe('count()', function () {
        it('should count up by value-funciton', function () {
            assert.strictEqual(_.count([1, 2, 3, 4, 5], function (e) {
                return e % 2 === 0;
            }), 2);
        });

        it('should return zero if empty', function () {
            assert.strictEqual(_.count([], function (e) {
                return e % 2 === 0;
            }), 0);
        });

        it('should return zero if null', function () {
            assert.strictEqual(_.count(null, function (e) {
                return e % 2 === 0;
            }), 0);
        });

        it('should return zero if undefined', function () {
            assert.strictEqual(_.count(undefined, function (e) {
                return e % 2 === 0;
            }), 0);
        });
    });

    describe('sum()', function () {
        it('should sum up by value-funciton', function () {
            assert.strictEqual(_.sum([1, 2, 3, 4, 5], function (e) {
                return e * 2;
            }), 30);
        });

        it('should return zero if empty', function () {
            assert.strictEqual(_.sum([], function (e) {
                return e * 2;
            }), 0);
        });

        it('should return zero if null', function () {
            assert.strictEqual(_.sum(null, function (e) {
                return e * 2;
            }), 0);
        });

        it('should return zero if undefined', function () {
            assert.strictEqual(_.sum(undefined, function (e) {
                return e * 2;
            }), 0);
        });
    });

    describe('consistOf()', function () {

        it('should return true if having single string property', function () {
            var testObj = { field1: 'value1', field2: 'value2', field3: 'value3' };
            assert.isOk(_.consistOf(testObj, 'field1'));
        });

        it('should return true if all elements have single string property', function () {
            var testArray = [{ field1: 'value1', field2: 'value2', field3: 'value3' }];
            assert.isOk(_.consistOf(testArray, 'field1'));
        });

        it('should return true if all elements have single string property', function () {
            var testArray = [{ field1: 'value1', field2: 'value2', field3: 'value3' }, { field1: 'value1', field4: 'value4', field5: 'value5' }];
            assert.isOk(_.consistOf(testArray, 'field1'));
        });

        it('should return true if having single string property as null', function () {
            var testObj = { field1: null, field2: 'value2', field3: 'value3' };
            assert.isOk(_.consistOf(testObj, 'field1'));
        });

        it('should return true if all elements have single string property as null', function () {
            var testArray = [{ field1: null, field2: 'value2', field3: 'value3' }];
            assert.isOk(_.consistOf(testArray, 'field1'));
        });

        it('should return true if all elements have string property as null', function () {
            var testArray = [{ field1: null, field2: 'value2', field3: 'value3' }, { field1: null, field4: 'value4', field5: 'value5' }];
            assert.isOk(_.consistOf(testArray, 'field1'));
        });

        it('should return true if having single string property as undefined', function () {
            var testObj = { field1: undefined, field2: 'value2', field3: 'value3' };
            assert.isOk(_.consistOf(testObj, 'field1'));
        });

        it('should return true if all elements have single string property as undefined', function () {
            var testArray = [{ field1: undefined, field2: 'value2', field3: 'value3' }];
            assert.isOk(_.consistOf(testArray, 'field1'));
        });

        it('should return true if all elements have single string property as undefined', function () {
            var testArray = [{ field1: undefined, field2: 'value2', field3: 'value3' }, { field1: undefined, field4: 'value4', field5: 'value5' }];
            assert.isOk(_.consistOf(testArray, 'field1'));
        });

        it('should return false if not having single string property', function () {
            var testObj = { field1: 'value1', field2: 'value2', field3: 'value3' };
            assert.isNotOk(_.consistOf(testObj, 'invalidField'));
        });

        it('should return false if element doesn\'t have single string property', function () {
            var testArray = [{ field1: 'value1', field2: 'value2', field3: 'value3' }];
            assert.isNotOk(_.consistOf(testArray, 'invalidField'));
        });

        it('should return false if element doesn\'t have single string property', function () {
            var testArray = [{ field1: 'value1', field2: 'value2', field3: 'value3' }, { field1: 'value1', field2: 'value2', field3: 'value3' }];
            assert.isNotOk(_.consistOf(testArray, 'invalidField'));
        });

        it('should return false if element doesn\'t have single string property', function () {
            var testArray = [{ field1: 'value1', field2: 'value2', field3: 'value3', fieldX: 'valueX' }, { field1: 'value1', field2: 'value2', field3: 'value3' }];
            assert.isNotOk(_.consistOf(testArray, 'fieldX'));
        });

        it('should return false if element doesn\'t have single string property', function () {
            var testArray = [{ field1: 'value1', field2: 'value2', field3: 'value3' }, { field1: 'value1', field2: 'value2', field3: 'value3', fieldX: 'valueX' }];
            assert.isNotOk(_.consistOf(testArray, 'fieldX'));
        });

        it('should return true if element have all properies', function () {
            var testObj = { field1: 'value1', field2: 'value2', field3: 'value3' };
            assert.isOk(_.consistOf(testObj, ['field1', 'field2', 'field3']));
        });

        it('should return true if all elements have all properies', function () {
            var testArray = [{ field1: 'value1', field2: 'value2', field3: 'value3' }];
            assert.isOk(_.consistOf(testArray, ['field1', 'field2', 'field3']));
        });

        it('should return true if all elements have all properies', function () {
            var testArray = [{ field1: 'value1', field2: 'value2', field3: 'value3' }, { field1: 'value1', field2: 'value2', field3: 'value3' }];
            assert.isOk(_.consistOf(testArray, ['field1', 'field2', 'field3']));
        });

        it('should return false if not having at least one', function () {
            var testObj = { field1: 'value1', field2: 'value2', field3: 'value3' };
            assert.isNotOk(_.consistOf(testObj, ['field1', 'field2', 'field3', 'invalidField']));
        });

        it('should return false if element doesn\'t have at least one', function () {
            var testArray = [{ field1: 'value1', field2: 'value2', field3: 'value3' }];
            assert.isNotOk(_.consistOf(testArray, ['field1', 'field2', 'field3', 'invalidField']));
        });

        it('should return false if element doesn\'t have at least one', function () {
            var testArray = [{ field1: 'value1', field2: 'value2', field3: 'value3' }, { field1: 'value1', field2: 'value2', field3: 'value3' }];
            assert.isNotOk(_.consistOf(testArray, ['field1', 'field2', 'field3', 'invalidField']));
        });

        it('should return true if having nested property', function () {
            var testObj = {
                field1: { field2: 'value2' }
            };
            assert.isOk(_.consistOf(testObj, { field1: 'field2' }));
        });

        it('should return true if element doesn\'t have nested property', function () {
            var testArray = [{
                field1: { field2: 'value2' }
            }];
            assert.isOk(_.consistOf(testArray, { field1: 'field2' }));
        });

        it('should return true if element doesn\'t have nested property', function () {
            var testArray = [{ field1: { field2: 'value2' } }, { field1: { field2: 'value2' } }];
            assert.isOk(_.consistOf(testArray, { field1: 'field2' }));
        });

        it('should return true if having nested property in array', function () {
            var testObj = {
                field1: { field2: 'value2' }
            };
            assert.isOk(_.consistOf(testObj, [{ field1: 'field2' }]));
        });

        it('should return true if element has nested property in array', function () {
            var testArray = [{
                field1: { field2: 'value2' }
            }];
            assert.isOk(_.consistOf(testArray, [{ field1: 'field2' }]));
        });

        it('should return true if element has nested property in array', function () {
            var testArray = [{ field1: { field2: 'value2' } }, { field1: { field2: 'value2' } }];
            assert.isOk(_.consistOf(testArray, [{ field1: 'field2' }]));
        });

        it('should return true if having nested property in array', function () {
            var testObj = {
                field1: 'value1',
                field2: { field3: 'value3' }
            };
            assert.isOk(_.consistOf(testObj, ['field1', { field2: 'field3' }]));
        });

        it('should return true if element has nested property in array', function () {
            var testArray = [{
                field1: 'value1',
                field2: { field3: 'value3' }
            }];
            assert.isOk(_.consistOf(testArray, ['field1', { field2: 'field3' }]));
        });

        it('should return true if element has nested property in array', function () {
            var testArray = [{ field1: 'value1', field2: { field3: 'value3' } }, { field1: 'value1', field2: { field3: 'value3' } }];
            assert.isOk(_.consistOf(testArray, ['field1', { field2: 'field3' }]));
        });

        it('should return true if having nested property in array', function () {
            var testObj = {
                field1: 'value1',
                field2: {
                    field3: 'value3',
                    field4: 'value4',
                    field5: 'value5'
                }
            };
            assert.isOk(_.consistOf(testObj, ['field1', { field2: ['field3', 'field4', 'field5'] }]));
        });

        it('should return true if element has nested property in array', function () {
            var testArray = [{
                field1: 'value1',
                field2: {
                    field3: 'value3',
                    field4: 'value4',
                    field5: 'value5'
                }
            }];
            assert.isOk(_.consistOf(testArray, ['field1', { field2: ['field3', 'field4', 'field5'] }]));
        });

        it('should return true if element has nested property in array', function () {
            var testArray = [{
                field1: 'value1',
                field2: {
                    field3: 'value3',
                    field4: 'value4',
                    field5: 'value5'
                }
            }, {
                field1: 'value1',
                field2: {
                    field3: 'value3',
                    field4: 'value4',
                    field5: 'value5'
                }
            }];
            assert.isOk(_.consistOf(testArray, ['field1', { field2: ['field3', 'field4', 'field5'] }]));
        });
    });
});