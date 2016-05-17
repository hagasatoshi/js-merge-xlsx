const Promise = require('bluebird');
const _ = require('underscore');
const JSZip = require('jszip');
const Mustache = require('mustache');
require('./lib/underscore_mixin');

const Excel = require('./lib/Excel');
const WorkBookXml = require('./lib/WorkBookXml');
const WorkBookRels = require('./lib/WorkBookRels');
const SheetXmls = require('./lib/SheetXmls');
const SharedStrings = require('./lib/SharedStrings');

const isNode = require('detect-node');
const config = {
    compression:        'DEFLATE',
    buffer_type_output: (isNode ? 'nodebuffer' : 'blob'),
    buffer_type_jszip:  (isNode ? 'nodebuffer' : 'arraybuffer')
};

const EXCEL_FILE = {
    SHARED_STRINGS:      'xl/sharedStrings.xml',
    WORKBOOK_RELS:       'xl/_rels/workbook.xml.rels',
    WORKBOOK:            'xl/workbook.xml',
    DIR_WORKSHEETS:      'xl/worksheets',
    DIR_WORKSHEETS_RELS: 'xl/worksheets/_rels'
};

const merge = (template, data, oututType = config.buffer_type_output) => {
    let templateObj = new JSZip(template);
    return templateObj.file(
        EXCEL_FILE.SHARED_STRINGS,
        Mustache.render(templateObj.file(EXCEL_FILE.SHARED_STRINGS).asText(), data)
    )
    .generate({type: oututType, compression: config.compression});
};

const bulkMergeToFiles = (template, arrayObj) => {
    return _.reduce(arrayObj, (zip, {name, data}) => {
        zip.file(name, merge(template, data, config.buffer_type_jszip));
        return zip;
    }, new JSZip())
    .generate({type: config.buffer_type_output, compression: config.compression});
};

const bulkMergeToSheets = (template, arrayObj) => {
    return parse(template)
    .then((templateObj) => {
        let excelObj = new Merge(templateObj)
            .addMergedSheets(arrayObj)
            .deleteTemplateSheet()
            .value();
        return new Excel(template).generateWithData(
            excelObj,
            {type: config.buffer_type_output, compression: config.compression}
        );
    });
};

const parse = (template) => {
    return new Excel(template).setTemplateSheetRel()
        .then(() => {
            return Promise.props({
                sharedstrings:   excel.parseSharedStrings(),
                workbookxmlRels: excel.parseWorkbookRels(),
                workbookxml:     excel.parseWorkbook(),
                sheetXmls:       excel.parseWorksheetsDir()
            })
        }).then(({sharedstrings, workbookxmlRels, workbookxml, sheetXmls}) => {
            let sheetXmlObjs = new SheetXmls(sheetXmls);
            return {
                relationship: new WorkBookRels(workbookxmlRels),
                workbookxml: new WorkBookXml(workbookxml),
                sheetXmls: sheetXmlObjs,
                templateSheetModel: sheetXmlObjs.getTemplateSheetModel(),
                sharedstrings: new SharedStrings(
                    sharedstrings, sheetXmlObjs.templateSheetData()
                )
            };
        });
};

class Merge {

    constructor(templateObj) {
        this.excelObj = templateObj;
    }

    addMergedSheets(dataArray) {
        this.excelObj = _.each(dataArray, ({newSheetName, mergeData}) => this.addMergedSheet(newSheetName, mergeData));
        return this;
    }

    addMergedSheet(newSheetName, mergeData) {
        let nextId = this.excelObj.relationship.nextRelationshipId();
        this.excelObj.relationship.add(nextId);
        this.excelObj.workbookxml.add(newSheetName, nextId);
        this.excelObj.sheetXmls.add(
            `sheet${nextId}.xml`,
            this.excelObj.templateSheetModel.cloneWithMergedString(
                this.excelObj.sharedstrings.addMergedStrings(mergeData)
            )
        );
    };

    deleteTemplateSheet() {
        let sheetname = this.workbookxml.firstSheetName();
        let targetSheet = this.findSheetByName(sheetname);
        this.relationship.delete(targetSheet.path);
        this.workbookxml.delete(sheetname);

        this.sheetXmls.delete(targetSheet.value.name);
        this.excel.removeWorksheet(targetSheet.value.name);
        return this;
    }

    findSheetByName(sheetname) {
        let sheetid = this.workbookxml.findSheetId(sheetname);
        if(!sheetid) {
            return null;
        }
        let targetFilePath = this.relationship.findSheetPath(sheetid);
        let targetFileName = _.last(targetFilePath.split('/'));
        return {path: targetFilePath, value: this.sheetXmls.find(targetFileName)};
    }

    value() {
        return this.excelObj;
    }
}

module.exports = {merge, bulkMergeToFiles, bulkMergeToSheets};
