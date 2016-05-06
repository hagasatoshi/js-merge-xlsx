/**
 * SheetModel.js
 * Manage each 'xl/worksheets/*.xml'
 * @author Satoshi Haga
 * @date 2016/05/06
 */

const _ = require('underscore');

class SheetModel {

    constructor(sheetObj) {
        this.sheetObj = sheetObj;
    }

    stringCount() {
        if(!this.sheetObj.worksheet) {
            return 0;
        }
        return _.reduce(
            this.sheetObj.worksheet.sheetData[0].row,
            (count, row) => {
                _.each(row.c, (cell) => {
                    if(cell['$'].t) {
                        count++;
                    }
                });
                return count;
            }, 0
        );
    }

    setStringIndex(stringIndex, cellAddress) {
        _.each(this.sheetObj.worksheet.sheetData[0].row, (row) => {
            _.each(row.c, (cell) => {
                if(cell['$'].r === cellAddress) {
                    cell.v[0] = stringIndex;
                }
            });
        });
    }

    value() {
        return this.sheetObj;
    }

    getName() {
        return this.sheetObj.name;
    }
}

module.exports = SheetModel;