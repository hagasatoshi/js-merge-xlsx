/**
 * SheetXmls.js
 * Manage 'xl/worksheets/*.xml'
 * @author Satoshi Haga
 * @date 2016/04/03
 */

const _ = require('underscore');
const SheetModel = require('./SheetModel');

class SheetXmls {

    constructor(sheetObjs) {
        this.sheetModels = _.map(sheetObjs, (sheetObj) => new SheetModel(sheetObj));
    }

    value() {
        return _.map(this.sheetModels, (sheetModel) => {
            if(!sheetModel.getName()) {
                return null;
            }
            let wrk = {};
            wrk.worksheet = {};
            _.extend(wrk.worksheet, sheetModel.value().worksheet);
            return {name: sheetModel.getName(), data: wrk};
        });
    }

    names() {
        return _.map(this.sheetModels, (sheetModel) => sheetModel.getName());
    }

    add(sheetId, sheetObj) {
        sheetObj.name = `sheet${sheetId}.xml`;
        this.sheetModels.push(new SheetModel(sheetObj));
    }

    delete(fileName) {
        _.splice(
            this.sheetModels,
            (sheetModel) => (sheetModel.value() && (sheetModel.getName() === fileName))
        );
    }

    find(fileName) {
        return _.find(
            this.sheetModels,
            (sheetModel) => sheetModel.value() && (sheetModel.getName() === fileName)
        );
    }

    stringCount() {
        return _.sum(this.sheetModels, (sheetModel) => sheetModel.stringCount());
    }

    templateSheetData() {
        return _.find(this.sheetModels, (sheetModel) => {
            return (sheetModel.getName().indexOf('.rels') === -1);
        }).value().worksheet.sheetData[0].row;
    }
}

module.exports = SheetXmls;
