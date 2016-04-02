/**
 * WorkBookRels.js
 * Manage 'xl/_rels/workbook.xml.rels'
 * @author Satoshi Haga
 * @date 2016/04/03
 */

const _ = require('underscore');
const OPEN_XML_SCHEMA_DEFINITION = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';

class WorkBookRels {

    constructor(workBookRels) {
        this.workBookRels = workBookRels;
        this.sheetRelationships = workBookRels.Relationships.Relationship;
    }

    valueWorkBookRels() {
        this.workBookRels.Relationships.Relationship = this.sheetRelationships;
        return this.workBookRels;
    }

    addSheetRelationship(sheetId) {
        this.sheetRelationships.push(
            { '$':
                { Id: sheetId,
                    Type: OPEN_XML_SCHEMA_DEFINITION,
                    Target: `worksheets/sheet${sheetId}.xml`
                }
            }
        );
    }

    deleteSheetRelationship(sheetPath) {
        _.each(this.sheetRelationships, (sheet, index) => {
            if(sheet && (sheet['$'].Target === sheetPath)) {
                this.sheetRelationships.splice(index,1);
            }
        });

    }

    findSheetPath(sheetId) {

        return _.max(this.sheetRelationships, (e)=>(e['$'].Id === sheetId))['$'].Target;
    }

    nextRelationshipId(){
        return _.max(this.sheetRelationships, (e)=> Number(e['$'].Id.replace('rId','')));
    }

}

module.exports = WorkBookRels;