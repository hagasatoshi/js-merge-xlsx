'use strict';

var _ = require('underscore');
require('../lib/underscore');
var assert = require('chai').assert;

describe('underscore.js', function () {
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