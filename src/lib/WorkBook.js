/**
 * WorkBook
 * @author Satoshi Haga
 * @date 2016/03/27
 */

const _ = require('underscore');

class WorkBook {

    constructor(workbookxml) {
        this.workbookxml = workbookxml;
        this.sheetDefinitions = new SheetDefinitions(workbookxml.workbook.sheets[0].sheet);
    }

    valueWorkbookxml() {
        this.workbookxml.workbook.sheets[0].sheet = this.sheetDefinitions.getValue();
        return this.workbookxml;
    }

    addSheetDefinition(sheetName, sheetId) {
        this.sheetDefinitions.add(sheetName, sheetId);
    }

    deleteSheetDefinition(sheetName) {
        this.sheetDefinitions.delete(sheetName);
    }

    findSheetDefinition(sheetName) {
        return this.sheetDefinitions.find(sheetName);
    }

    firstSheetName() {
        return this.sheetDefinitions.firstSheetName();
    }
}

class SheetDefinitions {

    constructor(sheets) {
        this.sheets = sheets;
    }

    getValue(){
        return this.sheets;
    }

    add(sheetName, sheetId) {
        this.sheets.push(
            { '$':
                { name: sheetName,
                    sheetId: sheetId.replace('rId',''),
                    'r:id': sheetId
                }
            }
        );
    }

    delete(sheetName) {
        _.each(this.sheets, (sheet, index) => {
            if(sheet && (sheet['$'].name === sheetName)) {
                this.sheets.splice(index,1);
            }
        });
    }

    find(sheetName) {
        return _.find(this.sheets, (e)=> (e['$'].name === sheetName));
    }

    firstSheetName() {
        return this.sheets[0]['$'].name;
    }

}

module.exports = WorkBook;