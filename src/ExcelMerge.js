/**
 * ExcelMerge
 * @author Satoshi Haga
 * @date 2016/03/27
 */

const Promise = require('bluebird');
const _ = require('underscore');
const JSZip = require('jszip');
const Mustache = require('mustache');
const {
    Excel, WorkBookXml, WorkBookRels, SheetXmls,
    SharedStrings, Config, underscore
} = require('require-dir')('./lib');

const ExcelMerge = {

    merge: (template, data, oututType = Config.JSZIP_OPTION.BUFFER_TYPE_OUTPUT) => {
        let templateObj = new JSZip(template);
        return templateObj.file(
            Config.EXCEL_FILES.FILE_SHARED_STRINGS,
            Mustache.render(templateObj.file(Config.EXCEL_FILES.FILE_SHARED_STRINGS).asText(), data)
        )
        .generate({type: oututType, compression: Config.JSZIP_OPTION.COMPLESSION});
    },

    bulkMergeToFiles: (template, arrayObj) => {
        return _.reduce(arrayObj, (zip, {name, data}) => {
            zip.file(name, ExcelMerge.merge(template, data, Config.JSZIP_OPTION.buffer_type_jszip));
            return zip;
        }, new JSZip())
        .generate({
            type:        Config.JSZIP_OPTION.BUFFER_TYPE_OUTPUT,
            compression: Config.JSZIP_OPTION.COMPLESSION
        });
    },

    bulkMergeToSheets: (template, arrayObj) => {
        return parse(template)
        .then((templateObj) => {
            let excelObj = new Merge(templateObj)
                .addMergedSheets(arrayObj)
                .deleteTemplateSheet()
                .value();
            return new Excel(template).generateWithData(excelObj);
        });
    }
};

const parse = (template) => {
    let templateObj = new Excel(template);
    return Promise.props({
        sharedstrings:   templateObj.parseSharedStrings(),
        workbookxmlRels: templateObj.parseWorkbookRels(),
        workbookxml:     templateObj.parseWorkbook(),
        sheetXmls:       templateObj.parseWorksheetsDir()
    }).then(({sharedstrings, workbookxmlRels, workbookxml, sheetXmls}) => {
        let sheetXmlObjs = new SheetXmls(sheetXmls);
        return {
            relationship:       new WorkBookRels(workbookxmlRels),
            workbookxml:        new WorkBookXml(workbookxml),
            sheetXmls:          sheetXmlObjs,
            templateSheetModel: sheetXmlObjs.getTemplateSheetModel(),
            sharedstrings:      new SharedStrings(
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
        _.each(dataArray, ({name, data}) => this.addMergedSheet(name, data));
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
        let sheetname = this.excelObj.workbookxml.firstSheetName();
        let targetSheet = this.findSheetByName(sheetname);
        this.excelObj.relationship.delete(targetSheet.path);
        this.excelObj.workbookxml.delete(sheetname);
        return this;
    }

    findSheetByName(sheetname) {
        let sheetid = this.excelObj.workbookxml.findSheetId(sheetname);
        if(!sheetid) {
            return null;
        }
        let targetFilePath = this.excelObj.relationship.findSheetPath(sheetid);
        let targetFileName = _.last(targetFilePath.split('/'));
        return {path: targetFilePath, value: this.excelObj.sheetXmls.find(targetFileName)};
    }

    value() {
        return this.excelObj;
    }
}

module.exports = ExcelMerge;
