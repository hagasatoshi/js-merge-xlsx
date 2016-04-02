/**
 * WorkBook
 * @author Satoshi Haga
 * @date 2016/03/27
 */

const _ = require('underscore');

class WorkBook {

    constructor(workbookxml) {
        this.workbookxml = workbookxml;
        this.sheetDefinitions = workbookxml.workbook.sheets[0].sheet;
    }

    valueWorkbookxml() {
        this.workbookxml.workbook.sheets[0].sheet = this.sheetDefinitions;
        return this.workbookxml;
    }

    addSheetDefinition(sheetName, sheetId) {
        this.sheetDefinitions.push(
            { '$':
                { name: sheetName,
                    sheetId: sheetId.replace('rId',''),
                    'r:id': sheetId
                }
            }
        );
    }

    deleteSheetDefinition(sheetName) {
        _.each(this.sheetDefinitions, (sheet,index)=>{
            if(sheet && (sheet['$'].name === sheetName)) {
                this.sheetDefinitions.splice(index,1);
            }
        });
    }

    findSheetDefinition(sheetName) {
        return _.find(this.sheetDefinitions, (e)=> (e['$'].name === sheetName));
    }

    firstSheetName() {
        return this.sheetDefinitions[0]['$'].name;
    }
}

module.exports = WorkBook;