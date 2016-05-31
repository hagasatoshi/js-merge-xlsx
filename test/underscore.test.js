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

    describe('variables()', function () {
        it('should parse from word surrounded by triple-brace', function () {
            var parsed = _.variables('{{{word}}}');
            assert.isOk(_.isArray(parsed));
            assert.strictEqual(parsed.length, 1);
            assert.strictEqual(parsed[0], 'word');
        });

        it('should parse from all word surrounded by triple-brace', function () {
            var parsed = _.variables('{{{word1}}}, {{{word2}}}}, {{{word3}}}');
            assert.isOk(_.isArray(parsed));
            assert.strictEqual(parsed.length, 3);
            assert.strictEqual(parsed[0], 'word1');
            assert.strictEqual(parsed[1], 'word2');
            assert.strictEqual(parsed[2], 'word3');
        });

        it('should not parse from word surrounded by double-brace', function () {
            var parsed = _.variables('{{word1}}');
            assert.isOk(_.isArray(parsed));
            assert.strictEqual(parsed.length, 0);
        });

        it('should not parse from word surrounded by double-brace', function () {
            var parsed = _.variables('{{word1}}, {{word2}}, {{word3}}');
            assert.isOk(_.isArray(parsed));
            assert.strictEqual(parsed.length, 0);
        });

        it('should not encode when parsing', function () {
            var parsed = _.variables('{{{<>\"\\\&\'}}}');
            assert.isOk(_.isArray(parsed));
            assert.strictEqual(parsed.length, 1);
            assert.strictEqual(parsed[0], '<>\"\\\&\'');
        });

        it('should return null if not string', function () {
            var parsed = _.variables(['{{{value1}}}']);
            assert.strictEqual(parsed, null);
        });

        it('should return null if null', function () {
            var parsed = _.variables(null);
            assert.strictEqual(parsed, null);
        });

        it('should return null if undefined', function () {
            var parsed = _.variables(undefined);
            assert.strictEqual(parsed, null);
        });
    });

    describe('hasVariable()', function () {
        it('should return true if having a triple-brace', function () {
            var hasVariable = _.hasVariable('{{{word}}}');
            assert.strictEqual(hasVariable, true);
        });

        it('should return true if having triple-braces', function () {
            var hasVariable = _.hasVariable('{{{word1}}}, {{{word2}}}');
            assert.strictEqual(hasVariable, true);
        });

        it('should return false if having a double-brace', function () {
            var hasVariable = _.hasVariable('{{word1}}');
            assert.strictEqual(hasVariable, false);
        });

        it('should return false if having double-braces', function () {
            var hasVariable = _.hasVariable('{{word1}}, {{word2}}');
            assert.strictEqual(hasVariable, false);
        });

        it('should return false if not string', function () {
            var hasVariable = _.hasVariable(['{{word1}}']);
            assert.strictEqual(hasVariable, false);
        });

        it('should return false if null', function () {
            var hasVariable = _.hasVariable(null);
            assert.strictEqual(hasVariable, false);
        });

        it('should return false if undefined', function () {
            var hasVariable = _.hasVariable(undefined);
            assert.strictEqual(hasVariable, false);
        });
    });

    describe('deepCopy()', function () {
        it('should throw error if string', function () {
            try {
                _.deepCopy('string');
                assert.isOk(false);
            } catch (e) {
                assert.isOk(true);
            }
        });

        it('should throw error if number', function () {
            try {
                _.deepCopy(10);
                assert.isOk(false);
            } catch (e) {
                assert.isOk(true);
            }
        });

        it('should throw error if boolean', function () {
            try {
                _.deepCopy(true);
                assert.isOk(false);
            } catch (e) {
                assert.isOk(true);
            }
        });

        it('should throw error if null', function () {
            try {
                _.deepCopy(null);
                assert.isOk(false);
            } catch (e) {
                assert.isOk(true);
            }
        });

        it('should throw error if undefined', function () {
            try {
                _.deepCopy(undefined);
                assert.isOk(false);
            } catch (e) {
                assert.isOk(true);
            }
        });

        it('should clone from object', function () {
            var cloned = _.deepCopy({ key1: 'value1', key2: 'value2' });
            assert.strictEqual(cloned.key1, 'value1');
            assert.strictEqual(cloned.key2, 'value2');
        });

        it('should clone from object including non-object value', function () {
            var cloned = _.deepCopy({ key1: 1, key2: true, key3: null, key4: undefined });
            assert.strictEqual(cloned.key1, 1);
            assert.strictEqual(cloned.key2, true);
            assert.strictEqual(cloned.key3, null);
            assert.strictEqual(cloned.key4, undefined);
        });

        it('should be the different reference when cloning object', function () {
            var source = { key1: 'value1', key2: 'value2' };
            var cloned = _.deepCopy(source);
            cloned.key1 = 'value3';
            assert.strictEqual(source.key1, 'value1');
        });

        it('should clone from array', function () {
            var cloned = _.deepCopy(['val1', 'val2', 'val3']);
            assert.strictEqual(_.isArray(cloned), true);
            assert.strictEqual(cloned.length, 3);
            assert.strictEqual(cloned[0], 'val1');
            assert.strictEqual(cloned[1], 'val2');
            assert.strictEqual(cloned[2], 'val3');
        });

        it('should be the different reference when cloning array', function () {
            var source = ['val1', 'val2', 'val3'];
            var cloned = _.deepCopy(source);
            cloned[0] = 'differentVal1';
            cloned[1] = 'differentVal2';
            cloned[2] = 'differentVal3';
            assert.strictEqual(source[0], 'val1');
            assert.strictEqual(source[1], 'val2');
            assert.strictEqual(source[2], 'val3');
        });

        it('should clone from array including non-object value', function () {
            var cloned = _.deepCopy([1, true, null]);
            assert.strictEqual(_.isArray(cloned), true);
            assert.strictEqual(cloned.length, 3);
            assert.strictEqual(cloned[0], 1);
            assert.strictEqual(cloned[1], true);
            assert.strictEqual(cloned[2], null);
        });

        it('should return null if undefined in array', function () {
            var cloned = _.deepCopy([undefined]);
            assert.strictEqual(_.isArray(cloned), true);
            assert.strictEqual(cloned.length, 1);
            assert.strictEqual(cloned[0], null);
        });

        it('should return undefined if undefined in object', function () {
            var cloned = _.deepCopy({ key: undefined });
            assert.strictEqual(cloned.key, undefined);
        });

        it('should clone from nested object', function () {
            var source = ['value1', { key2: 'value2' }, 1, true, null, undefined, { key3: 1, key4: true, key5: null, key6: undefined }];
            var cloned = _.deepCopy(source);
            assert.strictEqual(cloned[0], 'value1');
            assert.strictEqual(cloned[1].key2, 'value2');
            assert.strictEqual(cloned[2], 1);
            assert.strictEqual(cloned[3], true);
            assert.strictEqual(cloned[4], null);
            assert.strictEqual(cloned[5], null);
            assert.strictEqual(cloned[6].key3, 1);
            assert.strictEqual(cloned[6].key4, true);
            assert.strictEqual(cloned[6].key5, null);
            assert.strictEqual(cloned[6].key6, undefined);
        });
    });

    describe('deleteProps()', function () {
        it('should delete property', function () {
            var deleted = _.deleteProps({ key1: 'value1', key2: 'value2' }, ['key1']);
            assert.strictEqual(deleted.key1, undefined);
            assert.strictEqual(deleted.key2, 'value2');
        });

        it('should delete property even if last property', function () {
            var deleted = _.deleteProps({ key1: 'value1' }, ['key1']);
            assert.strictEqual(deleted.key1, undefined);
            assert.notStrictEqual(deleted, null);
            assert.notStrictEqual(deleted, undefined);
        });

        it('should not fail if invalid property name', function () {
            var deleted = _.deleteProps({ key1: 'value1' }, ['invalidKey']);
            assert.strictEqual(deleted.key1, 'value1');
        });

        it('should delete all properties', function () {
            var deleted = _.deleteProps({ key1: 'value1', key2: 'value2', key3: 'value3' }, ['key1', 'key3']);
            assert.strictEqual(deleted.key1, undefined);
            assert.strictEqual(deleted.key2, 'value2');
            assert.strictEqual(deleted.key3, undefined);
        });

        it('should delete property of all elements', function () {
            var target = [{ key1: 'value11', key2: 'value21' }, { key1: 'value12', key2: 'value22' }, { key1: 'value13', key2: 'value23' }];
            var deleted = _.deleteProps(target, ['key1']);
            assert.strictEqual(target[0].key1, undefined);
            assert.strictEqual(target[0].key2, 'value21');
            assert.strictEqual(target[1].key1, undefined);
            assert.strictEqual(target[1].key2, 'value22');
            assert.strictEqual(target[2].key1, undefined);
            assert.strictEqual(target[2].key2, 'value23');
        });

        it('should delete property of all elements', function () {
            var target = [{ key1: 'value11', key2: 'value21' }, { key3: 'value12', key4: 'value22' }, { key5: 'value13', key6: 'value23' }];
            var deleted = _.deleteProps(target, ['key1']);
            assert.strictEqual(target[0].key1, undefined);
            assert.strictEqual(target[0].key2, 'value21');
            assert.strictEqual(target[1].key3, 'value12');
            assert.strictEqual(target[1].key4, 'value22');
            assert.strictEqual(target[2].key5, 'value13');
            assert.strictEqual(target[2].key6, 'value23');
        });
    });

    describe('reverseEach()', function () {
        it('should call element of array in reverse', function () {
            var test = '';
            _.reverseEach(['first', 'second', 'third'], function (e) {
                test += e + '/';
            });
            assert.strictEqual(test, 'third/second/first/');
        });
    });

    describe('reduceInReverse()', function () {
        it('should call element of array in reverse', function () {
            var test = _.reduceInReverse(['first', 'second', 'third'], function (combined, e) {
                return combined + '/' + e;
            }, '');
            assert.strictEqual(test, '/third/second/first');
        });
    });

    describe('nestedEach()', function () {
        var appendString = function appendString(array1, array2) {
            var ret = '';
            _.nestedEach(array1, array2, function (e1, e2) {
                ret += e1 + '-' + e2 + '/';
            });
            return ret;
        };

        it('should call each element of both arrays', function () {
            var appended = appendString(['a', 'b', 'c'], ['1', '2', '3']);
            assert.strictEqual(appended, 'a-1/a-2/a-3/b-1/b-2/b-3/c-1/c-2/c-3/');
        });

        it('should not call if array1 is empty', function () {
            var appended = appendString([], ['1', '2', '3']);
            assert.strictEqual(appended, '');

            appended = appendString(null, ['1', '2', '3']);
            assert.strictEqual(appended, '');

            appended = appendString(undefined, ['1', '2', '3']);
            assert.strictEqual(appended, '');
        });

        it('should not call if array2 is empty', function () {
            var appended = appendString(['1', '2', '3'], []);
            assert.strictEqual(appended, '');

            appended = appendString(['1', '2', '3'], null);
            assert.strictEqual(appended, '');

            appended = appendString(['1', '2', '3'], undefined);
            assert.strictEqual(appended, '');
        });

        it('should not call if both array are empty', function () {
            var appended = appendString([], []);
            assert.strictEqual(appended, '');

            appended = appendString(null, null);
            assert.strictEqual(appended, '');

            appended = appendString(undefined, undefined);
            assert.strictEqual(appended, '');
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