/**
 * SheetXmls.js
 * Manage 'xl/worksheets/*.xml'
 * @author Satoshi Haga
 * @date 2016/04/03
 */

const _ = require('underscore');

class SheetXmls {

    constructor(files) {
        this.files = files;
    }

    value() {
        return _.map(this.files, (file) => {
            if(!file.name) {
                return null;
            }
            let sheetObj = {};
            sheetObj.worksheet = {};
            _.extend(sheetObj.worksheet, file.worksheet);
            return {name: file.name, data: sheetObj};
        });
    }

    names() {
        return _.map(this.files, (file) => file.name);
    }

    add(sheetId, file) {
        file.name = `sheet${sheetId}.xml`;
        this.files.push(file);
    }

    delete(fileName) {
        _.each(this.files, (file, index) => {
            if(file && (file.name === fileName)) {
                this.files.splice(index, 1);
            }
        });
    }

    find(fileName) {
        return _.find(this.files, (e) => (e.name === fileName));
    }

    stringCount() {
        let stringCount = 0;
        _.each(this.files, (sheet) => {
            if(sheet.worksheet) {
                _.each(sheet.worksheet.sheetData[0].row, (row) => {
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
        return _.find(this.files, (e) => {
            return (e.name.indexOf('.rels') === -1);
        }).worksheet.sheetData[0].row;
    }
}

module.exports = SheetXmls;
