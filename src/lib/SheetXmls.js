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
        this.sheetModels = _.reduce(
            sheetObjs,
            (models, sheetObj) => {
                models.push(new SheetModel(sheetObj));
                return models;
            },
            []
        )
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
        _.each(this.sheetModels, (sheetModel, index) => {
            if(sheetModel.value() && (sheetModel.getName() === fileName)) {
                this.sheetModels.splice(index, 1);
            }
        });
    }

    find(fileName) {
        return _.find(this.sheetModels, (sheetModel) => {
            if(!sheetModel.value()) {
                return false;
            }
            return (sheetModel.getName() === fileName);
        });
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
