const _ = require('underscore');
require('../lib/underscore');
const assert = require('chai').assert;

describe('underscore.js', () => {
    describe('consistOf()', () => {

        it('should return true if element have all properies', () => {
            let testObj = {field1: 'value1', field2: 'value2', field3: 'value3'};
            assert.isOk(_.consistOf(testObj, ['field1', 'field2', 'field3']));
        });
    });
});