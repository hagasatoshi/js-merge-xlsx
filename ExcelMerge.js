"use strict";function _classCallCheck(e,r){if(!(e instanceof r))throw new TypeError("Cannot call a class as a function")}var _createClass=function(){function e(e,r){for(var t=0;t<r.length;t++){var n=r[t];n.enumerable=n.enumerable||!1,n.configurable=!0,"value"in n&&(n.writable=!0),Object.defineProperty(e,n.key,n)}}return function(r,t,n){return t&&e(r.prototype,t),n&&e(r,n),r}}(),Promise=require("bluebird"),_=require("underscore"),JSZip=require("jszip"),Mustache=require("mustache"),_require=require("require-dir")("./lib"),Excel=_require.Excel,WorkBookXml=_require.WorkBookXml,WorkBookRels=_require.WorkBookRels,SheetXmls=_require.SheetXmls,SharedStrings=_require.SharedStrings,Config=_require.Config,underscore_mixin=_require.underscore_mixin,ExcelMerge={merge:function(e,r){var t=arguments.length<=2||void 0===arguments[2]?Config.JSZIP_OPTION.BUFFER_TYPE_OUTPUT:arguments[2],n=new JSZip(e);return n.file(Config.EXCEL_FILES.FILE_SHARED_STRINGS,Mustache.render(n.file(Config.EXCEL_FILES.FILE_SHARED_STRINGS).asText(),r)).generate({type:t,compression:Config.JSZIP_OPTION.COMPLESSION})},bulkMergeToFiles:function(e,r){return _.reduce(r,function(r,t){var n=t.name,i=t.data;return r.file(n,ExcelMerge.merge(e,i,Config.JSZIP_OPTION.buffer_type_jszip)),r},new JSZip).generate({type:Config.JSZIP_OPTION.BUFFER_TYPE_OUTPUT,compression:Config.JSZIP_OPTION.COMPLESSION})},bulkMergeToSheets:function(e,r){return parse(e).then(function(t){var n=new Merge(t).addMergedSheets(r).value();return new Excel(e).generateWithData(n)})}},parse=function(e){return new Excel(e).setTemplateSheetRel().then(function(e){return Promise.props({sharedstrings:e.parseSharedStrings(),workbookxmlRels:e.parseWorkbookRels(),workbookxml:e.parseWorkbook(),sheetXmls:e.parseWorksheetsDir()})}).then(function(e){var r=e.sharedstrings,t=e.workbookxmlRels,n=e.workbookxml,i=e.sheetXmls,o=new SheetXmls(i);return{relationship:new WorkBookRels(t),workbookxml:new WorkBookXml(n),sheetXmls:o,templateSheetModel:o.getTemplateSheetModel(),sharedstrings:new SharedStrings(r,o.templateSheetData())}})},Merge=function(){function e(r){_classCallCheck(this,e),this.excelObj=r}return _createClass(e,[{key:"addMergedSheets",value:function(e){var r=this;return _.each(e,function(e){var t=e.name,n=e.data;return r.addMergedSheet(t,n)}),this}},{key:"addMergedSheet",value:function(e,r){var t=this.excelObj.relationship.nextRelationshipId();this.excelObj.relationship.add(t),this.excelObj.workbookxml.add(e,t),this.excelObj.sheetXmls.add("sheet"+t+".xml",this.excelObj.templateSheetModel.cloneWithMergedString(this.excelObj.sharedstrings.addMergedStrings(r)))}},{key:"deleteTemplateSheet",value:function(){var e=this.excelObj.workbookxml.firstSheetName(),r=this.findSheetByName(e);return this.excelObj.relationship["delete"](r.path),this.excelObj.workbookxml["delete"](e),this}},{key:"findSheetByName",value:function(e){var r=this.excelObj.workbookxml.findSheetId(e);if(!r)return null;var t=this.excelObj.relationship.findSheetPath(r),n=_.last(t.split("/"));return{path:t,value:this.excelObj.sheetXmls.find(n)}}},{key:"value",value:function(){return this.excelObj}}]),e}();module.exports=ExcelMerge;