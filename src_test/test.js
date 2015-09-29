/**
 * * test.js
 * * Test script for js-merge-xlsx
 * * @author Satoshi Haga
 * * @date 2015/09/30
 **/

import assert from 'assert';
import ExcelMerge from '../excelmerge';

describe('sampleTest',  ()=>{
    describe('sample', ()=>{
        it('this is sample src_test', ()=>{
            var merge = new ExcelMerge('1');
            assert.equal(merge.excel, '1');
        });
    });
});