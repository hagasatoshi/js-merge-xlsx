/**
 * SheetXmls.js
 * Manage 'xl/worksheets/*.xml'
 * @author Satoshi Haga
 * @date 2016/04/03
 */

const _ = require('underscore');

class SheetXmls {

    constructor(sheetObjs) {
        this.sheetObjs = sheetObjs;
    }

    value() {
        return _.map(this.sheetObjs, (sheetObj) => {
            if(!sheetObj.name) {
                return null;
            }
            let sheetModel = {};
            sheetModel.worksheet = {};
            _.extend(sheetModel.worksheet, sheetObj.worksheet);
            return {name: sheetObj.name, data: sheetModel};
        });
    }

    names() {
        return _.map(this.sheetObjs, (sheetObj) => sheetObj.name);
    }

    add(sheetId, sheetObj) {
        sheetObj.name = `sheet${sheetId}.xml`;
        this.sheetObjs.push(sheetObj);
    }

    delete(fileName) {
        _.each(this.sheetObjs, (sheetObj, index) => {
            if(sheetObj && (sheetObj.name === fileName)) {
                this.sheetObjs.splice(index, 1);
            }
        });
    }

    find(fileName) {
        return _.find(this.sheetObjs, (sheetObj) => (sheetObj.name === fileName));
    }

    stringCount() {
        let stringCount = 0;
        _.each(this.sheetObjs, (sheetObj) => {
            if(sheetObj.worksheet) {
                _.each(sheetObj.worksheet.sheetData[0].row, (row) => {
                    _.each(row.c, (cell) => {
                        if(cell['$'].t) {
                            stringCount++;
                        }
                    });
                });
            }
        });
        return stringCount;
    }

    templateSheetData() {
        return _.find(this.sheetObjs, (sheetObj) => {
            return (sheetObj.name.indexOf('.rels') === -1);
        }).worksheet.sheetData[0].row;
    }
}

module.exports = SheetXmls;
