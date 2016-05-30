/**
 * SharedStrings.js
 * Manage 'xl/sharedStrings.xml'
 * @author Satoshi Haga
 * @date 2016/04/03
 */

const _ = require('underscore');
require('./underscore');
const Mustache = require('mustache');

class SharedStrings {

    constructor(rawObj, templateSheetObj) {
        this.rawObj = rawObj;
        this.templateSheetData = templateSheetObj;
        this.stringModels = this.parseStringModels();
    }

    parseStringModels() {
        if(!this.rawObj.sst || !this.rawObj.sst.si) {
            return null;
        }
        let stringModels = _.deepCopy(this.rawObj.sst.si);
        _.each(this.filterStaticString(stringModels), (stringModel) => {
            stringModel.usingCells = _.reduce(
                this.templateSheetData,
                (usingCells, row) => {
                    return usingCells.concat(
                        SharedStrings.usingCellAddresses(row.c, stringModel.sharedIndex)
                    );
                }, []
            );
        });
        return stringModels;
    }

    static usingCellAddresses(cells, stringIndex) {
        return _.reduce(cells, (addresses, cell) => {
            if(cell['$'].t === 's' && stringIndex === (cell.v[0] >> 0)) {
                addresses.push(cell['$'].r);
            }
            return addresses;
        }, []);
    }

    add(newStringModels) {
        let currentCount = this.stringModels.length;
        _.each(newStringModels, (e, index) => {
            e.sharedIndex = currentCount + index;
        });
        this.stringModels = this.stringModels.concat(newStringModels);
        return newStringModels;
    }

    value() {
        if(!this.stringModels) {
            return null;
        }
        this.rawObj.sst.si = _.deleteProperties(this.stringModels, ['sharedIndex', 'usingCells']);
        this.rawObj.sst['$'].uniqueCount = this.stringModels.length;
        this.rawObj.sst['$'].count = this.stringModels.length;
        return this.rawObj;
    }

    filterStaticString(stringModels) {
        let ret = [];
        _.each(stringModels, (stringObj, index) => {
            if(_.stringValue(stringObj.t) && _.hasVariable(_.stringValue(stringObj.t))) {
                stringObj.sharedIndex = index;
                ret.push(stringObj);
            }
        });
        return ret;
    }

    hasString() {
        return !!this.stringModels;
    }

    buildNewSharedStrings(mergedData) {
        return _.reduce(
            _.deepCopy(this.filterStaticString(this.stringModels)),
            (newSharedStrings, templateString) => {
                templateString.t[0] = Mustache.render(_.stringValue(templateString.t), mergedData);
                newSharedStrings.push(templateString);
                return newSharedStrings;
            }
        , []);
    }

    addMergedStrings(mergedData) {
        if(!this.hasString()) {
            return;
        }
        return this.add(
            this.buildNewSharedStrings(mergedData)
        );
    }
}

module.exports = SharedStrings;
