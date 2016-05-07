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
            _.sum(this.getRows(), stringCountInRow) :
            0 ;
    }

    value() {
        return this.sheetObj;
    }

    getName() {
        return this.sheetObj.name;
    }

    getRows() {
        return this.sheetObj.worksheet.sheetData[0].row;
    }

    clone(sheetSelected = false) {
        let cloned = new SheetModel(_.deepCopy(this.sheetObj));
        cloned.setSheetSelected(sheetSelected);
        return cloned;
    }

    setSheetSelected(sheetSelected) {
        this.sheetObj.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected =
            sheetSelected ? '1' : '0';
    }

    updateStringIndex(stringModels) {
        const updateIndex = (stringModel, row) => {
            _.each(stringModel.usingCells, (address) => {
                let cellAtThisAddress = _.find(row.c, (cell) => (cell['$'].r === address));
                if(cellAtThisAddress) {
                    cellAtThisAddress.v[0] = stringModel.sharedIndex;
                }
            });
        };
        _.nestedEach(stringModels, this.getRows(), updateIndex);
        return this;
    }

    cloneWithMergedString(stringModels) {
        return this.clone().updateStringIndex(stringModels);
    }
}

module.exports = SheetModel;