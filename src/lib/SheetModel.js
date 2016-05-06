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
        const stringCountInRow = (cells) => _.count(cells, (cell) => !!cell['$'].t);
        return this.sheetObj.worksheet ?
            _.sum(this.sheetObj.worksheet.sheetData[0].row, stringCountInRow) :
            0 ;
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