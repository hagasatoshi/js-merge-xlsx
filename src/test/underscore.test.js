const _ = require('underscore');
require('../lib/underscore');
const assert = require('chai').assert;

describe('underscore.js', () => {

    describe('stringValue()', () => {
        it('should return the same value if not array', () => {
            assert.strictEqual(_.stringValue('test'), 'test');
        });
        it('should return first element if array', () => {
            assert.strictEqual(_.stringValue(['first', 'second', 'third']), 'first');
        });
        it('should return attribute "mustache" if have', () => {
            assert.strictEqual(_.stringValue([{_: 'mustache', key1: 'value1'}]), 'mustache');
        });
    });

    describe('variables()', () => {
        it('should parse from word surrounded by triple-brace', () => {
            let parsed = _.variables('{{{word}}}');
            assert.isOk(_.isArray(parsed));
            assert.strictEqual(parsed.length, 1);
            assert.strictEqual(parsed[0], 'word');
        });

        it('should parse from all word surrounded by triple-brace', () => {
            let parsed = _.variables('{{{word1}}}, {{{word2}}}}, {{{word3}}}');
            assert.isOk(_.isArray(parsed));
            assert.strictEqual(parsed.length, 3);
            assert.strictEqual(parsed[0], 'word1');
            assert.strictEqual(parsed[1], 'word2');
            assert.strictEqual(parsed[2], 'word3');
        });

        it('should not parse from word surrounded by double-brace', () => {
            let parsed = _.variables('{{word1}}');
            assert.isOk(_.isArray(parsed));
            assert.strictEqual(parsed.length, 0);
        });

        it('should not parse from word surrounded by double-brace', () => {
            let parsed = _.variables('{{word1}}, {{word2}}, {{word3}}');
            assert.isOk(_.isArray(parsed));
            assert.strictEqual(parsed.length, 0);
        });

        it('should not encode when parsing', () => {
            let parsed = _.variables('{{{<>\"\\\&\'}}}');
            assert.isOk(_.isArray(parsed));
            assert.strictEqual(parsed.length, 1);
            assert.strictEqual(parsed[0], '<>\"\\\&\'');
        });

        it('should return null if not string', () => {
            let parsed = _.variables(['{{{value1}}}']);
            assert.strictEqual(parsed, null);
        });

        it('should return null if null', () => {
            let parsed = _.variables(null);
            assert.strictEqual(parsed, null);
        });

        it('should return null if undefined', () => {
            let parsed = _.variables(undefined);
            assert.strictEqual(parsed, null);
        });

    });

    describe('hasVariable()', () => {
        it('should return true if having a triple-brace', () => {
            let hasVariable = _.hasVariable('{{{word}}}');
            assert.strictEqual(hasVariable, true);
        });

        it('should return true if having triple-braces', () => {
            let hasVariable = _.hasVariable('{{{word1}}}, {{{word2}}}');
            assert.strictEqual(hasVariable, true);
        });

        it('should return false if having a double-brace', () => {
            let hasVariable = _.hasVariable('{{word1}}');
            assert.strictEqual(hasVariable, false);
        });

        it('should return false if having double-braces', () => {
            let hasVariable = _.hasVariable('{{word1}}, {{word2}}');
            assert.strictEqual(hasVariable, false);
        });

        it('should return false if not string', () => {
            let hasVariable = _.hasVariable(['{{word1}}']);
            assert.strictEqual(hasVariable, false);
        });

        it('should return false if null', () => {
            let hasVariable = _.hasVariable(null);
            assert.strictEqual(hasVariable, false);
        });

        it('should return false if undefined', () => {
            let hasVariable = _.hasVariable(undefined);
            assert.strictEqual(hasVariable, false);
        });

    });

    describe('count()', () => {
        it('should count up by value-funciton', () => {
            assert.strictEqual(
                _.count([1, 2, 3, 4, 5], (e) => (e % 2 === 0)),
                2
            );
        });

        it('should return zero if empty', () => {
            assert.strictEqual(_.count([], (e) => (e % 2 === 0)), 0);
        });

        it('should return zero if null', () => {
            assert.strictEqual(_.count(null, (e) => (e % 2 === 0)), 0);
        });

        it('should return zero if undefined', () => {
            assert.strictEqual(_.count(undefined, (e) => (e % 2 === 0)), 0);
        });
    });

    describe('sum()', () => {
        it('should sum up by value-funciton', () => {
            assert.strictEqual(
                _.sum([1, 2, 3, 4, 5], (e) => e*2),
                30
            );
        });

        it('should return zero if empty', () => {
            assert.strictEqual(_.sum([], (e) => e*2), 0);
        });

        it('should return zero if null', () => {
            assert.strictEqual(_.sum(null, (e) => e*2), 0);
        });

        it('should return zero if undefined', () => {
            assert.strictEqual(_.sum(undefined, (e) => e*2), 0);
        });
    });

    describe('consistOf()', () => {

        it('should return true if having single string property', () => {
            let testObj = {field1: 'value1', field2: 'value2', field3: 'value3'};
            assert.isOk(_.consistOf(testObj, 'field1'));
        });

        it('should return true if all elements have single string property', () => {
            let testArray = [{field1: 'value1', field2: 'value2', field3: 'value3'}];
            assert.isOk(_.consistOf(testArray, 'field1'));
        });

        it('should return true if all elements have single string property', () => {
            let testArray = [
                {field1: 'value1', field2: 'value2', field3: 'value3'},
                {field1: 'value1', field4: 'value4', field5: 'value5'}
            ];
            assert.isOk(_.consistOf(testArray, 'field1'));
        });

        it('should return true if having single string property as null', () => {
            let testObj = {field1: null, field2: 'value2', field3: 'value3'};
            assert.isOk(_.consistOf(testObj, 'field1'));
        });

        it('should return true if all elements have single string property as null', () => {
            let testArray = [{field1: null, field2: 'value2', field3: 'value3'}];
            assert.isOk(_.consistOf(testArray, 'field1'));
        });

        it('should return true if all elements have string property as null', () => {
            let testArray = [
                {field1: null, field2: 'value2', field3: 'value3'},
                {field1: null, field4: 'value4', field5: 'value5'}
            ];
            assert.isOk(_.consistOf(testArray, 'field1'));
        });

        it('should return true if having single string property as undefined', () => {
            let testObj = {field1: undefined, field2: 'value2', field3: 'value3'};
            assert.isOk(_.consistOf(testObj, 'field1'));
        });

        it('should return true if all elements have single string property as undefined', () => {
            let testArray = [{field1: undefined, field2: 'value2', field3: 'value3'}];
            assert.isOk(_.consistOf(testArray, 'field1'));
        });

        it('should return true if all elements have single string property as undefined', () => {
            let testArray = [
                {field1: undefined, field2: 'value2', field3: 'value3'},
                {field1: undefined, field4: 'value4', field5: 'value5'}
            ];
            assert.isOk(_.consistOf(testArray, 'field1'));
        });

        it('should return false if not having single string property', () => {
            let testObj = {field1: 'value1', field2: 'value2', field3: 'value3'};
            assert.isNotOk(_.consistOf(testObj, 'invalidField'));
        });

        it('should return false if element doesn\'t have single string property', () => {
            let testArray = [{field1: 'value1', field2: 'value2', field3: 'value3'}];
            assert.isNotOk(_.consistOf(testArray, 'invalidField'));
        });

        it('should return false if element doesn\'t have single string property', () => {
            let testArray = [
                {field1: 'value1', field2: 'value2', field3: 'value3'},
                {field1: 'value1', field2: 'value2', field3: 'value3'}
            ];
            assert.isNotOk(_.consistOf(testArray, 'invalidField'));
        });

        it('should return false if element doesn\'t have single string property', () => {
            let testArray = [
                {field1: 'value1', field2: 'value2', field3: 'value3', fieldX: 'valueX'},
                {field1: 'value1', field2: 'value2', field3: 'value3'}
            ];
            assert.isNotOk(_.consistOf(testArray, 'fieldX'));
        });

        it('should return false if element doesn\'t have single string property', () => {
            let testArray = [
                {field1: 'value1', field2: 'value2', field3: 'value3'},
                {field1: 'value1', field2: 'value2', field3: 'value3', fieldX: 'valueX'}
            ];
            assert.isNotOk(_.consistOf(testArray, 'fieldX'));
        });

        it('should return true if element have all properies', () => {
            let testObj = {field1: 'value1', field2: 'value2', field3: 'value3'};
            assert.isOk(_.consistOf(testObj, ['field1', 'field2', 'field3']));
        });

        it('should return true if all elements have all properies', () => {
            let testArray = [{field1: 'value1', field2: 'value2', field3: 'value3'}];
            assert.isOk(_.consistOf(testArray, ['field1', 'field2', 'field3']));
        });

        it('should return true if all elements have all properies', () => {
            let testArray = [
                {field1: 'value1', field2: 'value2', field3: 'value3'},
                {field1: 'value1', field2: 'value2', field3: 'value3'}
            ];
            assert.isOk(_.consistOf(testArray, ['field1', 'field2', 'field3']));
        });

        it('should return false if not having at least one', () => {
            let testObj = {field1: 'value1', field2: 'value2', field3: 'value3'};
            assert.isNotOk(_.consistOf(testObj, ['field1', 'field2', 'field3', 'invalidField']));
        });

        it('should return false if element doesn\'t have at least one', () => {
            let testArray = [{field1: 'value1', field2: 'value2', field3: 'value3'}];
            assert.isNotOk(_.consistOf(testArray, ['field1', 'field2', 'field3', 'invalidField']));
        });

        it('should return false if element doesn\'t have at least one', () => {
            let testArray = [
                {field1: 'value1', field2: 'value2', field3: 'value3'},
                {field1: 'value1', field2: 'value2', field3: 'value3'}
            ];
            assert.isNotOk(_.consistOf(testArray, ['field1', 'field2', 'field3', 'invalidField']));
        });

        it('should return true if having nested property', () => {
            let testObj = {
                field1: {field2: 'value2'}
            };
            assert.isOk(_.consistOf(testObj, {field1: 'field2'}));
        });

        it('should return true if element doesn\'t have nested property', () => {
            let testArray = [{
                field1: {field2: 'value2'}
            }];
            assert.isOk(_.consistOf(testArray, {field1: 'field2'}));
        });

        it('should return true if element doesn\'t have nested property', () => {
            let testArray = [
                {field1: {field2: 'value2'}},
                {field1: {field2: 'value2'}}
            ];
            assert.isOk(_.consistOf(testArray, {field1: 'field2'}));
        });

        it('should return true if having nested property in array', () => {
            let testObj = {
                field1: {field2: 'value2'}
            };
            assert.isOk(_.consistOf(testObj, [{field1: 'field2'}]));
        });

        it('should return true if element has nested property in array', () => {
            let testArray = [{
                field1: {field2: 'value2'}
            }];
            assert.isOk(_.consistOf(testArray, [{field1: 'field2'}]));
        });

        it('should return true if element has nested property in array', () => {
            let testArray = [
                {field1: {field2: 'value2'}},
                {field1: {field2: 'value2'}}
            ];
            assert.isOk(_.consistOf(testArray, [{field1: 'field2'}]));
        });

        it('should return true if having nested property in array', () => {
            let testObj = {
                field1: 'value1',
                field2: {field3: 'value3'}
            };
            assert.isOk(_.consistOf(testObj, ['field1', {field2: 'field3'}]));
        });

        it('should return true if element has nested property in array', () => {
            let testArray = [{
                field1: 'value1',
                field2: {field3: 'value3'}
            }];
            assert.isOk(_.consistOf(testArray, ['field1', {field2: 'field3'}]));
        });

        it('should return true if element has nested property in array', () => {
            let testArray = [
                {field1: 'value1', field2: {field3: 'value3'}},
                {field1: 'value1', field2: {field3: 'value3'}}
            ];
            assert.isOk(_.consistOf(testArray, ['field1', {field2: 'field3'}]));
        });

        it('should return true if having nested property in array', () => {
            let testObj = {
                field1: 'value1',
                field2: {
                    field3: 'value3',
                    field4: 'value4',
                    field5: 'value5'
                }
            };
            assert.isOk(_.consistOf(testObj, ['field1', {field2: ['field3', 'field4', 'field5']}]));
        });

        it('should return true if element has nested property in array', () => {
            let testArray = [{
                field1: 'value1',
                field2: {
                    field3: 'value3',
                    field4: 'value4',
                    field5: 'value5'
                }
            }];
            assert.isOk(
                _.consistOf(testArray, ['field1', {field2: ['field3', 'field4', 'field5']}])
            );
        });

        it('should return true if element has nested property in array', () => {
            let testArray = [
                {
                    field1: 'value1',
                    field2: {
                        field3: 'value3',
                        field4: 'value4',
                        field5: 'value5'
                    }
                },
                {
                    field1: 'value1',
                    field2: {
                        field3: 'value3',
                        field4: 'value4',
                        field5: 'value5'
                    }
                }
            ];
            assert.isOk(
                _.consistOf(testArray, ['field1', {field2: ['field3', 'field4', 'field5']}])
            );
        });

    });
});