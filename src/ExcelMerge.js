/**
 * * ExcelMerge
 * * Template managing class. wrapping JsZip object.
 * * @author Satoshi Haga
 * * @date 2015/09/30
 **/

import Mustache from 'mustache'

class ExcelMerge{

    /**
     * * constructor
     * * @param {Object} excel JsZip object including MS-Excel file
     **/
    constructor(excel){
        this.excel = excel;
    }

    /**
     * * render
     * * @param {Object} bind_data binding data
     * * @param {Object} jszip_option JsZip#generate() option.
     * * @returns {Object} rendered MS-Excel data. data-format is determined by jszip_option
     **/
    render(bind_data, jszip_option={type: "blob",compression:"DEFLATE"}){
        let template = this.excel.file('xl/sharedStrings.xml').asText();
        this.excel.file('xl/sharedStrings.xml', Mustache.render(template, bind_data));
        return this.excel.generate(jszip_option);
    }
}

//Exports
module.exports = ExcelMerge;