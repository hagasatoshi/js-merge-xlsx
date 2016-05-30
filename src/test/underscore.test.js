const _ = require('underscore');
require('../lib/underscore');
const assert = require('chai').assert;

describe('underscore.js', () => {
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