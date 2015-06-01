###
  JavaScript excel template engine
  @author Satoshi Haga(satoshi.haga.github@gmail.com)
###
  
#require
Promise = require('bluebird');
xml2js = require('xml2js');
parseString = Promise.promisify(xml2js.parseString);
builder = new xml2js.Builder();
JSZip = require("jszip");
_ = require("underscore");
Log = require('log');
log = new Log(Log.DEBUG);


class spreadsheet  
  ###
    Constructor
    Read
    @param {arraybuffer} excel file data(openXML format)
  ###
  constructor: (excel)->
    @zip = new JSZip(excel);
    Promise.props
      shared_strings: parseString @zip.file('xl/sharedStrings.xml').asText(),
      workbookxml_rels: parseString @zip.file('xl/_rels/workbook.xml.rels').asText(),
      workbookxml: parseString @zip.file('xl/workbook.xml').asText(),
      sheet_xmls :@_parse_dir_in_excel('xl/worksheets')
    .then (template_obj)=>
      @shared_strings.initialize template_obj.shared_strings;
      @workbookxml_rels = template_obj.workbookxml_rels;
      @workbookxml = template_obj.workbookxml;
      @sheet_xmls = template_obj.sheet_xmls;
      log.info 'SpreadSheet is initialized successfully'

  ###
    Return excel data
    @param {String} blob/ arraybuffer / nodebuffer
    @return {Object} excel data formatted by parameter 'generate_type'
  ###
  generate: (generate_type)=>
    log.info 'SpreadSheet:generate'
    @zip
    .file "xl/_rels/workbook.xml.rels",builder.buildObject(@workbookxml_rels)
    .file "xl/workbook.xml",builder.buildObject(@workbookxml)
    .file 'xl/sharedStrings.xml', builder.buildObject(@shared_strings.get_obj())
    _.each @sheet_xmls, (sheet)=>
      if sheet.name
        sheet_obj = {}
        sheet_obj.worksheet = {}
        _.extend sheet_obj.worksheet, sheet.worksheet
        @zip.file "xl/worksheets/#{sheet.name}", builder.buildObject(sheet_obj)

    @zip.generate type:generate_type

  ###
    Activate the specific sheet of this file
    @param {String} name of target sheet
  ###
  active_sheet: (sheetname)->
    target_sheet_name = @sheet_by_name sheetname
    _.each @sheet_xmls, (sheet)=>
      if sheet.name == target_sheet_name.value.worksheet.name
        sheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected = '1'
      else
        sheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected = '0'

  ###
    Delete the specific sheet of this file
    @param {String} sheet name
  ###
  delete_sheet: (sheetname)=>
    target_sheet = @sheet_by_name sheetname
    for sheet,index in @workbookxml_rels.Relationships.Relationship
      if sheet['$'].Target == target_sheet.path
        @workbookxml_rels.Relationships.Relationship.splice index,1
    for sheet,index in @workbookxml.workbook.sheets[0].sheet
      if sheet['$'].name == sheetname
        @workbookxml.workbook.sheets[0].sheet.splice index,1
    for sheet_xml,index in @sheet_xmls
      if sheet_xml.name == target_sheet.value.name
        @sheet_xmls.splice index,1

  copy_sheet: (src_sheet_name,dest_sheet_name)=>
    next_id = @available_sheetid();
    @workbookxml_rels.Relationships.Relationship.push { 
      '$': { 
        Id: next_id,
        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        Target: "worksheets/sheet#{next_id}.xml"
      }
    }
    @workbookxml.workbook.sheets[0].sheet.push { 
      '$': { 
        name: dest_sheet_name, 
        sheetId: next_id.replace('rId',''), 
        'r:id': next_id 
      } 
    };

    src_sheet = @sheet_by_name(src_sheet_name).value;
    copied_sheet = JSON.parse(JSON.stringify(src_sheet));
    copied_sheet.name = "sheet#{next_id}.xml";
    @sheet_xmls.push copied_sheet


  bulk_copy_sheet: (src_sheet_name,dest_sheet_names)=>
    _.each dest_sheet_names, (dest_sheet_name)->
      @copy_sheet src_sheet_name, dest_sheet_name

  available_sheetid: ()=>
    max_rel = _.max @workbookxml_rels.Relationships.Relationship,(e)-> Number e['$'].Id.replace('rId','')
    next_id = 'rId' + ('00' + (parseInt((max_rel['$'].Id.replace('rId','')))+parseInt(1))).slice(-3);

  sheet_by_name: (sheetname)=>
    sheetid = _.find(@workbookxml.workbook.sheets[0].sheet, (e)-> e['$'].name == sheetname)['$']['r:id']
    target_file_path = _.max(@workbookxml_rels.Relationships.Relationship,(e)->e['$'].Id == sheetid)['$'].Target
    target_file_name = target_file_path.split('/')[target_file_path.split('/').length-1];
    sheet_xml = _.find this.sheet_xmls,(e)-> e.name == target_file_name
    sheet = path: target_file_path, value: sheet_xml


  cell_by_name: (sheetname,cell_name)=>
    cell_name_array = cell_name.split('');
    index = 0;
    _.each cell_name_array, (c)->
      if(/^[a-zA-Z()]+$/.test(c))
        index++
    row_string = cell_name.substr index, cell_name.length-index
    row = @row_by_name sheetname,row_string
    
    return undefined if row == undefined
        
    cell = _.find row.c, (c)-> c['$'].r == cell_name

  set_row: (sheetname,row_number,cell_values,existing_setting)=>
    key_values = []
    _.each cell_values, (cell_value,index)->
      col_string = _convert_alphabet(index)
      key_values.push {cell_name: (col_string+row_number), value:cell_value}
    @bulk_set_value sheetname, key_values,existing_setting
    
  bulk_set_value: (sheetname,key_values,existing_setting)->
    sheet_xml = @sheet_by_name sheetname
    _.each key_values, (cell)->
        @set_value sheet_xml,sheetname,cell.cell_name,cell.value,existing_setting


  set_value: (sheet_xml,sheetname,cell_name,value,existing_setting)=>
    return if !value
    cell_value = {}
    if _is_number(value)
      cell_value = { '$': { r: cell_name }, v: [ value ] }
    else
      next_index =
        if existing_setting && existing_setting[value]
          existing_setting[value]
        else
          @shared_strings.add_string(value)
      cell_value = { '$': { r: cell_name, t: 's' }, v: [ next_index ] }

    row_string = _get_row_string(cell_name);
    row = @row_by_name(sheetname,row_string);
    cell = @cell_by_name(sheetname, cell_name);
    if row == undefined
      new_row = {
        '$': { r: row_string, spans: '1:5' },
        c: [cell_value]
      }
      sheet_xml.value.worksheet.sheetData[0].row.push(new_row);
    else if cell == undefined
      row.c.push(cell_value)
      @_update_row(sheet_xml, row)
    else
      cell_value['$'].s = cell['$'].s
      if(cell_value['$'].t)
        cell['$'].t = 's'
        cell.v = [ next_index ];
      else
        cell.v = [ value ];
        if(cell['$'].t) 
          delete cell['$'].t
      
      @_update_cell(sheet_xml, row, cell);
    
      
  row_by_name: (sheetname,row_number)=>
    sheet_xml = @sheet_by_name(sheetname);
    row = _.find sheet_xml.value.worksheet.sheetData[0].row, (e)-> e['$'].r == row_number

  _update_row: (sheet,row)=>
    row.c = _.sortBy row.c, (e)-> _revert_number(_col_string(e['$'].r))
    _.each sheet.value.worksheet.sheetData[0].row, (existing_row)->
      if existing_row['$'].r == row['$'].r
        existing_row = row

  _update_cell: (sheet,row,cell)=>
    row.c = _.sortBy row.c, (e)-> _revert_number(_col_string(e['$'].r))
    _.each sheet.value.worksheet.sheetData[0].row, (existing_row)->
      if existing_row['$'].r == row['$'].r
        _.each existing_row.c, (existing_cell)->
          if(existing_cell['$'].r == cell['$'].r)
            existing_cell = cell;

  _parse_dir_in_excel: (dir)=>
    files = @zip.folder(dir).file(/.xml/);
    file_xmls = [];
    files.reduce (promise, file)->
      promise.then (prior_file)->
        Promise.resolve()
        .then ()->
          parseString @zip.file(file.name).asText()
        .then (file_xml)->
          file_xml.name = file.name.split('/')[file.name.split('/').length-1]
          file_xmls.push(file_xml)
      ,Promise.resolve()

  class shared_strings
    initialize: (obj)=>
      @obj = obj
      @count = parseInt(obj.sst.si.length)-parseInt(1)
      
    get_obj: ()->@obj

    add_string: (value)=>
      value = '' if !value
      new_string = t: [ value ], phoneticPr: ['$': { fontId: '1' } ]
      @obj.sst.si.push(new_string);
      @count = parseInt(this.obj.sst.si.length) - parseInt(1);

      
_is_number = (value)->
  if typeof(value) != 'number' && typeof(value) != 'string'
    false;
  else
    value == parseFloat(value) && isFinite(value)


_convert = (value)-> 
  'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('')[value]


_convert_alphabet = (value)->
  number1 = Math.floor value/(26*26);
  number2 = Math.floor (value-number1*26*26)/26;
  number3 = value-(number1*26*26+number2*26);
  alphabet1 = _convert(number1) == 'A' ? '' : _convert(number1 - 1)
  alphabet2 = (alphabet1 == '' && _convert(number2) == 'A') ? '' : _convert(number2 - 1);
  alphabet3 = _convert(number3)

  alphabet = alphabet1 + alphabet2 + alphabet3;


_revert = (alphabet)->
  alphabet.charCodeAt(0) - 'A'.charCodeAt(0) + 1

_revert_number = (alphabet)->
  alphabet_with_zero = ('00'+alphabet).slice(-3).split('')
  value = 
    if(alphabet_with_zero[0] != '0')
      value = value + _revert(alphabet_with_zero[0])*26*26;
    else if(alphabet_with_zero[1] != '0')
      value = value + _revert(alphabet_with_zero[1])*26;
    else if(alphabet_with_zero[2] != '0')
      value = value + _revert(alphabet_with_zero[2]);
    else
      0

_col_string = (cell_name)->
  cell_name_array = cell_name.split ''
  index = 0
  _.each cell_name_array, (c)->
    index++; if (/^[a-zA-Z()]+$/.test(c))
  col_string = cell_name.substr(0, index)


_get_row_string = (cell_name)->
  cell_name_array = cell_name.split ''
  index = 0
  _.each cell_name_array, (c)->
    index++ if (/^[a-zA-Z()]+$/.test(c))

  row_string = cell_name.substr(index, cell_name.length - index)

_get_col_string = (cell_name)->
  cell_name_array = cell_name.split ''
  index = 0
  _.each cell_name_array,(c)->
    index++ if (/^[a-zA-Z()]+$/.test(c))
  col_string = cell_name.substr(0, index)

load_config = ()->
  config = yaml.safeLoad(fs.readFileSync('./yaml/config.yml', 'utf8'))

module.exports = spreadsheet;

