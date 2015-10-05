/**
 * * SpreadSheet
 * * Manage MS-Excel file
 * * @author Satoshi Haga
 * * @date 2015/10/03
 **/
import Mustache from 'mustache'
import Promise from 'bluebird'
import _ from 'underscore'
import './underscore_mixin'
import JSZip from 'jszip'
import isNode from 'detect-node'
const output_buffer = {type: (isNode?'nodebuffer':'blob'), compression:"DEFLATE"};
const jszip_buffer = {type: (isNode?'nodebuffer':'arraybuffer'), compression:"DEFLATE"};
import xml2js from 'xml2js'
var parseString = Promise.promisify(xml2js.parseString);
var builder = new xml2js.Builder();


class SpreadSheet{

    /** member variables */
    //excel : {Object} JSZip instance including template excel file
    //variables : {Array} including mustache-variables defined in sharedstrings.xml
    //sharedstrings : {Array} includings common strings defined in sharedstrings.xml
    //sharedstrings_obj : {Object} whole sharedstrings object
    //sharedstrings_str : {String} whole sharedstrings string
    //common_strings_with_variable : {Array} including common strings only having mustache variables
    //sheet_xmls : {Array} including objects parsed from  'xl/worksheets/*.xml'
    //template_sheet_data : {Object} object parsed from 'xl/worksheets/*.xml'. this is used as template-file
    //template_sheet_name : {String} sheet-name of template-file
    //workbookxml_rels : {Object} parsed from 'xl/_rels/workbook.xml.rels'
    //workbookxml : {Object} parsed from 'xl/workbook.xml'

    /**
     * * load
     * * @param {Object} excel JsZip object including MS-Excel file
     * * @return {Promise|Object} Promise instance including this
     **/
    load(excel){

        //set member variable
        this.excel = excel;
        this.sharedstrings_str = excel.file('xl/sharedStrings.xml').asText();
        this.variables = _(this.sharedstrings_str).variables();
        this.common_strings_with_variable = [];

        //some members are parsed in promise-chain because xml2js parses asynchronously
        return Promise.props({
            sharedstrings_obj: parseString(this.sharedstrings_str),
            workbookxml_rels: parseString(this.excel.file('xl/_rels/workbook.xml.rels').asText()),
            workbookxml: parseString(this.excel.file('xl/workbook.xml').asText()),
            sheet_xmls :this._parse_dir_in_excel('xl/worksheets')
        }).then((templates)=>{
            this.sharedstrings_obj = templates.sharedstrings_obj;
            this.sharedstrings = templates.sharedstrings_obj.sst.si;
            this.workbookxml_rels = templates.workbookxml_rels;
            this.workbookxml = templates.workbookxml;
            this.sheet_xmls = templates.sheet_xmls;
            this.template_sheet_data = _.find(templates.sheet_xmls,(e)=>(e.name.indexOf('.rels') === -1)).worksheet.sheetData[0].row;
            this.template_sheet_name = this.workbookxml.workbook.sheets[0].sheet[0]['$'].name;
            this.common_strings_with_variable = this._parse_common_string_with_variable();

            //return this for chaining
            return this;
        });
    }

    /**
     * * simple_render
     * * @param {Object} bind_data binding data
     * * @returns {Promise|Object} rendered MS-Excel data. data-format is determined by jszip_option
     **/
    simple_render(bind_data){
        return Promise.resolve().then(()=>this._simple_render(bind_data, output_buffer));
    }

    /**
     * * bulk_render_multi_file
     * * @param {Array} bind_data_array including data{name: file's name, data: binding-data}
     * * @returns {Promise|Object} rendered MS-Excel data.
     **/
    bulk_render_multi_file(bind_data_array){

        var all_excels = new JSZip();
        _.each(bind_data_array, (bind_data)=>{
            all_excels.file(bind_data.name, this._simple_render(bind_data.data, jszip_buffer));
        });
        return Promise.resolve().then(()=> all_excels.generate(output_buffer));
    }

    /**
     * * add_sheet_binding_data
     * * @param {String} dest_sheet_name name of new sheet
     * * @param {Object} data binding data
     * * @return {Promise|Object} Excel data. format is determinated by parameter
     **/
    add_sheet_binding_data(dest_sheet_name, data){
        //1.add relation of next sheet
        let next_id = this._available_sheetid();
        this.workbookxml_rels.Relationships.Relationship.push(
            { '$':
            { Id: next_id,
                Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                Target: 'worksheets/sheet'+next_id+'.xml'
            }
            }
        );
        this.workbookxml.workbook.sheets[0].sheet.push({ '$': { name: dest_sheet_name, sheetId: next_id.replace('rId',''), 'r:id': next_id } });

        //2.add sheet file.
        //2-1.prepare rendered-strings
        let rendered_strings = JSON.parse(JSON.stringify(this.common_strings_with_variable));
        _.each(rendered_strings,(e)=>{
            e.t[0] = Mustache.render(_(e.t).string_value(), data);
        });

        //2-2.add rendered-string into sharedstrings
        let current_count = this.sharedstrings.length;
        _.each(rendered_strings,(e,index)=>{
            e.shared_index = current_count + index;
            this.sharedstrings.push(e);
        });

        //2-4.build new sheet oject
        let source_sheet = this._sheet_by_name(this.template_sheet_name).value;
        let added_sheet = this._build_new_sheet(source_sheet, rendered_strings);

        //2-5.update sheet name.
        added_sheet.name = 'sheet'+next_id+'.xml';

        //2-6.add this sheet into sheet_xmls
        this.sheet_xmls.push(added_sheet);
    }


        /**
     * * generate
     * * call JSZip#generate() binding current data
     * * @param {Object} option option for JsZip#genereate()
     * * @return {Object} Excel data. format is determinated by parameter
     **/
    generate(option){
        //sharedstrings
        this.sharedstrings_obj.sst.si = this.sharedstrings;
        this.sharedstrings_obj.sst['$'].count = this.sharedstrings_obj.sst['$'].uniqueCount = this.sharedstrings.length;
        this.excel.file('xl/sharedStrings.xml', builder.buildObject(this.sharedstrings_obj));
        //workbook.xml.rels
        this.excel.file("xl/_rels/workbook.xml.rels",builder.buildObject(this.workbookxml_rels));
        //workbook.xml
        this.excel.file("xl/workbook.xml",builder.buildObject(this.workbookxml));
        //sheet_xmls
        _.each(this.sheet_xmls, (sheet)=>{
            if(sheet.name){
                var sheet_obj = {};
                sheet_obj.worksheet = {};
                _.extend(sheet_obj.worksheet, sheet.worksheet);
                this.excel.file('xl/worksheets/'+sheet.name, builder.buildObject(sheet_obj));
            }
        });
        //call JSZip#generate()
        return this.excel.generate(option);
    }

    /**
     * * _simple_render
     * * @param {Object} bind_data binding data
     * * @param {Object} option JsZip#generate() option.
     * * @returns {Object} rendered MS-Excel data. data-format is determined by jszip_option
     * * @private
     **/
    _simple_render(bind_data, option){
        return this.excel
            .file('xl/sharedStrings.xml', Mustache.render(this.sharedstrings_str, bind_data))
            .generate(option);
    }

    /**
     * * _parse_common_string_with_variable
     * * @return {Array} including common strings only having mustache-variable
     * * @private
     **/
    _parse_common_string_with_variable(){

        let common_strings_with_variable = [];

        _.each(this.sharedstrings,(string_obj, index)=>{
            if(_(_(string_obj.t).string_value()).has_variable()){
                string_obj.shared_index = index;
                common_strings_with_variable.push(string_obj);
            }
        });
        _.each(common_strings_with_variable, (common_string_with_variable)=>{
            common_string_with_variable.using_cells = [];
            _.each(this.template_sheet_data,(row)=>{
                _.each(row.c,(cell)=>{
                    if(cell['$'].t === 's'){
                        if(common_string_with_variable.shared_index === parseInt(cell.v[0])){
                            common_string_with_variable.using_cells.push(cell['$'].r);
                        }
                    }
                });
            });
        });

        return common_strings_with_variable;
    }

    /**
     * * _parse_dir_in_excel
     * * @param {String} dir directory name in Zip file.
     * * @return {Promise|Array} array including files parsed by xml2js
     * * @private
     **/
    _parse_dir_in_excel(dir){
        let files = this.excel.folder(dir).file(/.xml/);
        let file_xmls = [];
        return files.reduce(
            (promise, file)=>
                promise.then((prior_file)=>
                    Promise.resolve()
                        .then(()=>parseString(this.excel.file(file.name).asText()))
                        .then((file_xml)=>{
                            file_xml.name = file.name.split('/')[file.name.split('/').length-1];
                            file_xmls.push(file_xml);
                            return file_xmls;
                        })
                )
            ,
            Promise.resolve()
        );
    }

    /**
     * * _build_new_sheet
     * * @param {Object} source_sheet
     * * @param {Array} common_strings_with_variable
     * * @return {Object}
     * * @private
     **/
    _build_new_sheet(source_sheet, common_strings_with_variable){
        let added_sheet = JSON.parse(JSON.stringify(source_sheet));
        _.each(common_strings_with_variable,(e,index)=>{
            _.each(e.using_cells, (cell_address)=>{
                _.each(added_sheet.worksheet.sheetData[0].row,(row)=>{
                    _.each(row.c,(cell)=>{
                        if(cell['$'].r === cell_address){
                            cell.v[0] = e.shared_index;
                        }
                    });
                });
            });
        });
        return added_sheet;
    }

    /**
     * * _available_sheetid
     * * @return {String} id of next sheet
     * * @private
     **/
    _available_sheetid(){
        let max_rel = _.max(this.workbookxml_rels.Relationships.Relationship, (e)=> Number(e['$'].Id.replace('rId','')));
        let next_id = 'rId' + ('00' + (parseInt((max_rel['$'].Id.replace('rId','')))+parseInt(1))).slice(-3);
        return next_id;
    }

    /**
     * * _sheet_by_name
     * * @param {String} sheetname target sheet name
     * * @return {Object} sheet object
     * * @private
     **/
    _sheet_by_name(sheetname){
        let sheetid = _.find(this.workbookxml.workbook.sheets[0].sheet, (e)=> (e['$'].name === sheetname))['$']['r:id'];
        let target_file_path = _.max(this.workbookxml_rels.Relationships.Relationship, (e)=>(e['$'].Id === sheetid))['$'].Target;
        let target_file_name = target_file_path.split('/')[target_file_path.split('/').length-1];
        let sheet_xml = _.find(this.sheet_xmls, (e)=>(e.name === target_file_name));
        let sheet = {path: target_file_path, value: sheet_xml};
        return sheet;
    }

}

//Exports
module.exports = SpreadSheet;