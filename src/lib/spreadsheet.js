/**
 * * SpreadSheet
 * * Manage MS-Excel file. core business-logic class for js-merge-xlsx.
 * * @author Satoshi Haga
 * * @date 2015/10/03
 **/
var Mustache = require('mustache');
var Promise = require('bluebird');
var _ = require('underscore');
require('./underscore_mixin');
var JSZip = require('jszip');
var isNode = require('detect-node');
const output_buffer = {type: (isNode?'nodebuffer':'blob'), compression:"DEFLATE"};
const jszip_buffer = {type: (isNode?'nodebuffer':'arraybuffer'), compression:"DEFLATE"};
var xml2js = require('xml2js');
var parseString = Promise.promisify(xml2js.parseString);
var builder = new xml2js.Builder();

const OPEN_XML_SCHEMA_DEFINITION = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';

class SpreadSheet{

    /**
     * * member variables
     * * excel {Object} JSZip instance including template excel file
     * * variables {Array} including mustache-variables defined in sharedstrings.xml
     * * sharedstrings {Array} includings common strings defined in sharedstrings.xml
     * * sharedstrings_obj {Object} whole sharedstrings object
     * * common_strings_with_variable {Array} including common strings only having mustache variables
     * * sheet_xmls {Array} including objects parsed from  'xl/worksheets/*.xml'
     * * sheet_xmls_rels {Array} including objects pared from 'xl/worksheets/_rels/*.xml.rels'
     * * template_sheet_data {Object} object parsed from 'xl/worksheets/*.xml'. this is used as template-file
     * * template_sheet_name {String} sheet-name of template-file
     * * workbookxml_rels {Object} parsed from 'xl/_rels/workbook.xml.rels'
     * * workbookxml {Object} parsed from 'xl/workbook.xml'
     * */


    /**
     * * load
     * * @param {Object} excel JsZip object including MS-Excel file
     * * @return {Promise|Object} Promise instance including this
     **/
    load(excel){

        //validation
        if(!(excel instanceof JSZip)) return Promise.reject('First parameter must be JSZip instance including MS-Excel data');

        //set member variable
        this.excel = excel;
        this.variables = _(excel.file('xl/sharedStrings.xml').asText()).variables();
        this.common_strings_with_variable = [];

        //some members are parsed in promise-chain because xml2js parses asynchronously
        return Promise.props({
            sharedstrings_obj: parseString(excel.file('xl/sharedStrings.xml').asText()),
            workbookxml_rels: parseString(this.excel.file('xl/_rels/workbook.xml.rels').asText()),
            workbookxml: parseString(this.excel.file('xl/workbook.xml').asText()),
            sheet_xmls: this._parse_dir_in_excel('xl/worksheets'),
            sheet_xmls_rels: this._parse_dir_in_excel('xl/worksheets/_rels')
        }).then((templates)=>{
            this.sharedstrings = templates.sharedstrings_obj.sst.si;
            this.workbookxml_rels = templates.workbookxml_rels;
            this.workbookxml = templates.workbookxml;
            this.sheet_xmls = templates.sheet_xmls;
            this.sheet_xmls_rels = templates.sheet_xmls_rels;
            this.template_sheet_name = this.workbookxml.workbook.sheets[0].sheet[0]['$'].name;
            this.template_sheet_data = _.find(templates.sheet_xmls,(e)=>(e.name.indexOf('.rels') === -1)).worksheet.sheetData[0].row;
            this.template_sheet_rels_data = _(this._template_sheet_rels()).deep_copy();
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

        //validation
        if(!bind_data) return Promise.reject('simple_render() must has parameter');

        return Promise.resolve().then(()=>this._simple_render(bind_data, output_buffer));
    }

    /**
     * * bulk_render_multi_file
     * * @param {Array} bind_data_array including data{name: file's name, data: binding-data}
     * * @returns {Promise|Object} rendered MS-Excel data.
     **/
    bulk_render_multi_file(bind_data_array){

        //validation
        if(!_.isArray(bind_data_array)) return Promise.reject('bulk_render_multi_file() has only array object');
        if(_.find(bind_data_array,(e)=>!(e.name && e.data))) return Promise.reject('bulk_render_multi_file() is called with invalid parameter');

        var all_excels = new JSZip();
        _.each(bind_data_array, ({name,data})=>all_excels.file(name, this._simple_render(data, jszip_buffer)));
        return Promise.resolve().then(()=> all_excels.generate(output_buffer));
    }

    /**
     * * add_sheet_binding_data
     * * @param {String} dest_sheet_name name of new sheet
     * * @param {Object} data binding data
     * * @return {Object} this instance for chaining
     **/
    add_sheet_binding_data(dest_sheet_name, data){

        //validation
        if((!dest_sheet_name) || !(data)) return Promise.reject('add_sheet_binding_data() needs to have 2 paramter.');

        //1.add relation of next sheet
        let next_id = this._available_sheetid();
        this.workbookxml_rels.Relationships.Relationship.push({ '$': { Id: next_id, Type: OPEN_XML_SCHEMA_DEFINITION, Target: 'worksheets/sheet'+next_id+'.xml'}});
        this.workbookxml.workbook.sheets[0].sheet.push({ '$': { name: dest_sheet_name, sheetId: next_id.replace('rId',''), 'r:id': next_id } });

        //2.add sheet file.
        //2-1.prepare rendered-strings
        let rendered_strings = _(this.common_strings_with_variable).deep_copy();
        _.each(rendered_strings,(e)=>e.t[0] = Mustache.render(_(e.t).string_value(), data));

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

        return this;
    }

    /**
     * * activate_sheet
     * * @param {String} sheetname target sheet name
     * * @return {Object} this instance for chaining
     **/
    activate_sheet(sheetname){

        //validation
        if(!sheetname) return Promise.reject('activate_sheet() needs to have 1 paramter.');

        let target_sheet_name = this._sheet_by_name(sheetname);
        if(!target_sheet_name) return Promise.reject(`Invalid sheet name '${sheetname}'.`);

        _.each(this.sheet_xmls, (sheet)=>{
            if(!sheet.worksheet) return;
            sheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected = (sheet.name === target_sheet_name.value.worksheet.name) ? '1' : '0';
        });
        return this;
    }

    /**
     * * forcus_on_first_sheet
     * * @return {Object} this instance for chaining
     **/
    forcus_on_first_sheet(){
        return this.activate_sheet(this._first_sheet_name());
    }

    /**
     * * delete_sheet
     * * @param {String} sheetname target sheet name
     * * @return {Object} this instance for chaining
     **/
    delete_sheet(sheetname){
        if(!sheetname) return Promise.reject('delete_sheet() needs to have 1 paramter.');

        let target_sheet = this._sheet_by_name(sheetname);
        if(!target_sheet) return Promise.reject(`Invalid sheet name '${sheetname}'.`);

        _.each(this.workbookxml_rels.Relationships.Relationship, (sheet,index)=>{
            if(sheet && (sheet['$'].Target === target_sheet.path)) {
                this.workbookxml_rels.Relationships.Relationship.splice(index,1);
            }
        });
        _.each(this.workbookxml.workbook.sheets[0].sheet, (sheet,index)=>{
            if(sheet && (sheet['$'].name === sheetname)){
                this.workbookxml.workbook.sheets[0].sheet.splice(index,1);
            }
        });
        _.each(this.sheet_xmls, (sheet_xml,index)=>{
            if(sheet_xml && (sheet_xml.name === target_sheet.value.name)){
                this.sheet_xmls.splice(index,1);
            }
        });
        return this;
    }

    /**
     * * delete_template_sheet
     * * @return {Object} this instance for chaining
     **/
    delete_template_sheet(){
        return this.delete_sheet(this.template_sheet_name);
    }

    /**
     * * has_as_shared_string
     * * @param {String} target_str
     * * @return {boolean}
     **/
    has_as_shared_string(target_str){
        return (this.excel.file('xl/sharedStrings.xml').asText().indexOf(target_str) !== -1)
    }

    /**
     * * generate
     * * call JSZip#generate() binding current data
     * * @param {Object} option option for JsZip#genereate()
     * * @return {Object} Excel data. format is determinated by parameter
     **/
    generate(option){
        parseString(this.excel.file('xl/sharedStrings.xml').asText())
        .then((sharedstrings_obj)=>{

            //sharedstring
            sharedstrings_obj.sst.si = this._clean_shared_strings();
            sharedstrings_obj.sst['$'].count = sharedstrings_obj.sst['$'].uniqueCount = this.sharedstrings.length;
            this.excel
                .file('xl/sharedStrings.xml', builder.buildObject(sharedstrings_obj))
                .file("xl/_rels/workbook.xml.rels",builder.buildObject(this.workbookxml_rels))
                .file("xl/workbook.xml",builder.buildObject(this.workbookxml));

            //sheet_xmls
            _.each(this.sheet_xmls, (sheet)=>{
                if(sheet.name){
                    var sheet_obj = {};
                    sheet_obj.worksheet = {};
                    _.extend(sheet_obj.worksheet, sheet.worksheet);
                    this.excel.file(`xl/worksheets/${sheet.name}`, builder.buildObject(sheet_obj));
                }
            });

            //sheet_xmls_rels
            let str_template_sheet_rels = builder.buildObject(this.template_sheet_rels_data);
            _.each(this.sheet_xmls, (sheet)=>{
                if(sheet.name){
                    this.excel.file(`xl/worksheets/_rels/${sheet.name}.rels`, str_template_sheet_rels);
                }
            });

            //call JSZip#generate()
            return this.excel.generate(option);
        })

    }


    /**
     * * _simple_render
     * * @param {Object} bind_data binding data
     * * @param {Object} option JsZip#generate() option.
     * * @returns {Object} rendered MS-Excel data. data-format is determined by jszip_option
     * * @private
     **/
    _simple_render(bind_data, option=output_buffer){
        return this.excel
            .file('xl/sharedStrings.xml', Mustache.render(this.excel.file('xl/sharedStrings.xml').asText(), bind_data))
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
            if(_(string_obj.t).string_value() && _(_(string_obj.t).string_value()).has_variable()){
                string_obj.shared_index = index;
                common_strings_with_variable.push(string_obj);
            }
        });
        _.each(common_strings_with_variable, (common_string_with_variable)=>{
            common_string_with_variable.using_cells = [];
            _.each(this.template_sheet_data,(row)=>{
                _.each(row.c,(cell)=>{
                    if(cell['$'].t === 's'){
                        if(common_string_with_variable.shared_index === (cell.v[0] >> 0)){
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
        let added_sheet = _(source_sheet).deep_copy();
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
        added_sheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected = '0';
        return added_sheet;
    }

    /**
     * * _available_sheetid
     * * @return {String} id of next sheet
     * * @private
     **/
    _available_sheetid(){
        let max_rel = _.max(this.workbookxml_rels.Relationships.Relationship, (e)=> Number(e['$'].Id.replace('rId','')));
        let next_id = 'rId' + ('00' + (((max_rel['$'].Id.replace('rId','') >> 0))+1)).slice(-3);
        return next_id;
    }

    /**
     * * _sheet_by_name
     * * @param {String} sheetname target sheet name
     * * @return {Object} sheet object
     * * @private
     **/
    _sheet_by_name(sheetname){
        let target_sheet = _.find(this.workbookxml.workbook.sheets[0].sheet, (e)=> (e['$'].name === sheetname));
        if(!target_sheet) return null;  //invalid sheet name

        let sheetid = target_sheet['$']['r:id'];
        let target_file_path = _.max(this.workbookxml_rels.Relationships.Relationship, (e)=>(e['$'].Id === sheetid))['$'].Target;
        let target_file_name = target_file_path.split('/')[target_file_path.split('/').length-1];
        let sheet_xml = _.find(this.sheet_xmls, (e)=>(e.name === target_file_name));
        let sheet = {path: target_file_path, value: sheet_xml};
        return sheet;
    }

    /**
     * * _sheet_rels_by_name
     * * @param {String} sheetname target sheet name
     * * @return {Object} sheet_rels object
     * * @private
     **/
    _sheet_rels_by_name(sheetname){
        let target_file_path = this._sheet_by_name(sheetname).path;
        let target_name = target_file_path.split('/')[target_file_path.split('/').length-1] + '.rels';
        let sheet_xml_rels = _.find(this.sheet_xmls_rels, (e)=>(e.name === target_name));
        let sheet = {name: target_name, value: sheet_xml_rels};
        return sheet;
    }

    /**
     * * _template_sheet_rels
     * * @return {Object} sheet_rels object of template-sheet
     * * @private
     **/
    _template_sheet_rels(){
        return this._sheet_rels_by_name(this.template_sheet_name);
    }


    /**
     * * _sheet_names
     * * @return {Array} array including sheet name
     * * @private
     **/
    _sheet_names(){
        return _.map(this.sheet_xmls, (e)=>e.name);
    }

    /**
     * * _clean_shared_strings
     * * @return {Array} shared strings
     * * @private
     **/
    _clean_shared_strings(){
        return _.map(this.sharedstrings, (e)=>{
            return {t:e.t, phoneticPr:e.phoneticPr}
        });
    }

    /**
     * * _first_sheet_name
     * * @return {String} name of first-sheet of MS-Excel file
     * * @private
     **/
    _first_sheet_name(){
        return this.workbookxml.workbook.sheets[0].sheet[0]['$'].name;
    }

    /**
     * * active_sheets
     * * @return {Array} array including only active sheets.
     * * @private
     **/
    _active_sheets(){
        return _.filter(this.sheet_xmls, (sheet)=>(sheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected === '1'));
    }

    /**
     * * deactive_sheets
     * * @return {Array} array including only deactive sheets.
     * * @private
     **/
    _deactive_sheets(){
        return _.filter(this.sheet_xmls, (sheet)=>(sheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected === '0'));
    }
}

//Exports
module.exports = SpreadSheet;
