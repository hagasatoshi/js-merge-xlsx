/**
 * SharedStrings.js
 * Manage 'xl/sharedStrings.xml'
 * @author Satoshi Haga
 * @date 2016/04/03
 */

const _ = require('underscore');
require('./underscore_mixin');

class SharedStrings {

    constructor(sharedstringsObj, templateSheetData) {
        this.rawData = sharedstringsObj;
        this.strings = sharedstringsObj.sst.si;

        this.setUsingCells(this.getOnlyHavingVariable(), templateSheetData);
    }

    setUsingCells(sharedStrings, templateSheetData) {
        _.each(sharedStrings, (str)=>{
            str.usingCells = [];
            _.each(templateSheetData, (row)=>{
                _.each(row.c,(cell)=>{
                    if(cell['$'].t === 's'){
                        if(str.sharedIndex === (cell.v[0] >> 0)){
                            str.usingCells.push(cell['$'].r);
                        }
                    }
                });
            });
        });
    }

    add(newStrings) {
        let currentCount = this.strings.length;
        _.each(newStrings, (e, index) => {
            e.sharedIndex = currentCount + index;
            this.strings.push(e);
        });
    }

    value() {
        if(!this.strings){
            return null;
        }
        this.rawData.sst.si = _.deleteProperties(this.strings, ['sharedIndex', 'usingCells']);
        this.rawData.sst['$'].uniqueCount = this.strings.length;
        this.rawData.sst['$'].count = this.strings.length;
        return this.rawData;
    }

    getOnlyHavingVariable() {
        let ret = [];
        _.each(this.strings, (stringObj, index)=>{
            if(_.stringValue(stringObj.t) && _.hasVariable(_.stringValue(stringObj.t))){
                stringObj.sharedIndex = index;
                ret.push(stringObj);
            }
        });
        return ret;
    }

    hasString(){
        return !!this.strings;
    }

}

module.exports = SharedStrings;
