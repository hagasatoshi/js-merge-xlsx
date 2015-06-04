
/*
  JavaScript excel template engine
  @author Satoshi Haga(satoshi.haga.github@gmail.com)
 */

(function() {
  var JSZip, Promise, SpreadSheet, _, _col_string, _convert, _convert_alphabet, _get_col_string, _get_row_string, _is_number, _revert, _revert_number, builder, load_config, parseString, xml2js,
    bind = function(fn, me){ return function(){ return fn.apply(me, arguments); }; };

  Promise = require('bluebird');

  xml2js = require('xml2js');

  parseString = Promise.promisify(xml2js.parseString);

  builder = new xml2js.Builder();

  JSZip = require("jszip");

  _ = require("underscore");

  SpreadSheet = (function() {
    var shared_strings;

    function SpreadSheet() {
      this._parse_dir_in_excel = bind(this._parse_dir_in_excel, this);
      this._update_cell = bind(this._update_cell, this);
      this._update_row = bind(this._update_row, this);
      this.row_by_name = bind(this.row_by_name, this);
      this.set_value = bind(this.set_value, this);
      this.set_row = bind(this.set_row, this);
      this.cell_by_name = bind(this.cell_by_name, this);
      this.sheet_by_name = bind(this.sheet_by_name, this);
      this.available_sheetid = bind(this.available_sheetid, this);
      this.bulk_copy_sheet = bind(this.bulk_copy_sheet, this);
      this.copy_sheet = bind(this.copy_sheet, this);
      this.delete_sheet = bind(this.delete_sheet, this);
      this.generate = bind(this.generate, this);
      this.initialize = bind(this.initialize, this);
    }


    /*
      Constructor
      Read
      @param {arraybuffer} excel file data(openXML format)
     */

    SpreadSheet.prototype.initialize = function(excel) {
      var spread;
      spread = this;
      this.zip = new JSZip(excel);
      return Promise.props({
        shared_strings: parseString(spread.zip.file('xl/sharedStrings.xml').asText()),
        workbookxml_rels: parseString(spread.zip.file('xl/_rels/workbook.xml.rels').asText()),
        workbookxml: parseString(spread.zip.file('xl/workbook.xml').asText()),
        sheet_xmls: spread._parse_dir_in_excel('xl/worksheets')
      }).then((function(_this) {
        return function(template_obj) {
          _this.shared_strings = new shared_strings(template_obj.shared_strings);
          _this.workbookxml_rels = template_obj.workbookxml_rels;
          _this.workbookxml = template_obj.workbookxml;
          _this.sheet_xmls = template_obj.sheet_xmls;
          return console.log('SpreadSheet is initialized successfully');
        };
      })(this));
    };


    /*
      Return excel data
      @param {String} blob/ arraybuffer / nodebuffer
      @return {Object} excel data formatted by parameter 'generate_type'
     */

    SpreadSheet.prototype.generate = function(generate_type) {
      console.log('SpreadSheet:generate');
      this.zip.file("xl/_rels/workbook.xml.rels", builder.buildObject(this.workbookxml_rels)).file("xl/workbook.xml", builder.buildObject(this.workbookxml)).file('xl/sharedStrings.xml', builder.buildObject(this.shared_strings.get_obj()));
      _.each(this.sheet_xmls, (function(_this) {
        return function(sheet) {
          var sheet_obj;
          if (sheet.name) {
            sheet_obj = {};
            sheet_obj.worksheet = {};
            _.extend(sheet_obj.worksheet, sheet.worksheet);
            return _this.zip.file("xl/worksheets/" + sheet.name, builder.buildObject(sheet_obj));
          }
        };
      })(this));
      return this.zip.generate({
        type: generate_type
      });
    };


    /*
      Activate the specific sheet of this file
      @param {String} name of target sheet
     */

    SpreadSheet.prototype.active_sheet = function(sheetname) {
      var target_sheet_name;
      target_sheet_name = this.sheet_by_name(sheetname);
      return _.each(this.sheet_xmls, (function(_this) {
        return function(sheet) {
          if (sheet.name === target_sheet_name.value.worksheet.name) {
            return sheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected = '1';
          } else {
            return sheet.worksheet.sheetViews[0].sheetView[0]['$'].tabSelected = '0';
          }
        };
      })(this));
    };


    /*
      Delete the specific sheet of this file
      @param {String} sheet name
     */

    SpreadSheet.prototype.delete_sheet = function(sheetname) {
      var target_sheet;
      target_sheet = this.sheet_by_name(sheetname);
      _.each(this.workbookxml_rels.Relationships.Relationship, (function(_this) {
        return function(sheet, index) {
          if (sheet['$'].Target === target_sheet.path) {
            return _this.workbookxml_rels.Relationships.Relationship.splice(index, 1);
          }
        };
      })(this));
      _.each(this.workbookxml.workbook.sheets[0].sheet, (function(_this) {
        return function(sheet, index) {
          if (sheet['$'].name === sheetname) {
            return _this.workbookxml.workbook.sheets[0].sheet.splice(index, 1);
          }
        };
      })(this));
      return _.each(this.sheet_xmls, (function(_this) {
        return function(sheet_xml, index) {
          if (sheet_xml.name === target_sheet.value.name) {
            return _this.sheet_xmls.splice(index, 1);
          }
        };
      })(this));
    };

    SpreadSheet.prototype.copy_sheet = function(src_sheet_name, dest_sheet_name) {
      var copied_sheet, next_id, src_sheet;
      next_id = this.available_sheetid();
      this.workbookxml_rels.Relationships.Relationship.push({
        '$': {
          Id: next_id,
          Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
          Target: "worksheets/sheet" + next_id + ".xml"
        }
      });
      this.workbookxml.workbook.sheets[0].sheet.push({
        '$': {
          name: dest_sheet_name,
          sheetId: next_id.replace('rId', ''),
          'r:id': next_id
        }
      });
      src_sheet = this.sheet_by_name(src_sheet_name).value;
      copied_sheet = JSON.parse(JSON.stringify(src_sheet));
      copied_sheet.name = "sheet" + next_id + ".xml";
      return this.sheet_xmls.push(copied_sheet);
    };

    SpreadSheet.prototype.bulk_copy_sheet = function(src_sheet_name, dest_sheet_names) {
      return _.each(dest_sheet_names, function(dest_sheet_name) {
        return this.copy_sheet(src_sheet_name, dest_sheet_name);
      });
    };

    SpreadSheet.prototype.available_sheetid = function() {
      var max_rel, next_id;
      max_rel = _.max(this.workbookxml_rels.Relationships.Relationship, function(e) {
        return Number(e['$'].Id.replace('rId', ''));
      });
      return next_id = 'rId' + ('00' + (parseInt(max_rel['$'].Id.replace('rId', '')) + parseInt(1))).slice(-3);
    };

    SpreadSheet.prototype.sheet_by_name = function(sheetname) {
      var sheet, sheet_xml, sheetid, target_file_name, target_file_path;
      sheetid = _.find(this.workbookxml.workbook.sheets[0].sheet, function(e) {
        return e['$'].name === sheetname;
      })['$']['r:id'];
      target_file_path = _.max(this.workbookxml_rels.Relationships.Relationship, function(e) {
        return e['$'].Id === sheetid;
      })['$'].Target;
      target_file_name = target_file_path.split('/')[target_file_path.split('/').length - 1];
      sheet_xml = _.find(this.sheet_xmls, function(e) {
        return e.name === target_file_name;
      });
      return sheet = {
        path: target_file_path,
        value: sheet_xml
      };
    };

    SpreadSheet.prototype.cell_by_name = function(sheetname, cell_name) {
      var cell, cell_name_array, index, row, row_string;
      cell_name_array = cell_name.split('');
      index = 0;
      _.each(cell_name_array, function(c) {
        if (/^[a-zA-Z()]+$/.test(c)) {
          return index++;
        }
      });
      row_string = cell_name.substr(index, cell_name.length - index);
      row = this.row_by_name(sheetname, row_string);
      if (row === void 0) {
        return void 0;
      }
      return cell = _.find(row.c, function(c) {
        return c['$'].r === cell_name;
      });
    };

    SpreadSheet.prototype.set_row = function(sheetname, row_number, cell_values, existing_setting) {
      var key_values;
      key_values = [];
      _.each(cell_values, function(cell_value, index) {
        var col_string;
        col_string = _convert_alphabet(index);
        return key_values.push({
          cell_name: col_string + row_number,
          value: cell_value
        });
      });
      return this.bulk_set_value(sheetname, key_values, existing_setting);
    };

    SpreadSheet.prototype.bulk_set_value = function(sheetname, key_values, existing_setting) {
      var sheet_xml;
      sheet_xml = this.sheet_by_name(sheetname);
      return _.each(key_values, (function(_this) {
        return function(cell) {
          return _this.set_value(sheet_xml, sheetname, cell.cell_name, cell.value, existing_setting);
        };
      })(this));
    };

    SpreadSheet.prototype.set_value = function(sheet_xml, sheetname, cell_name, value, existing_setting) {
      var cell, cell_value, new_row, next_index, row, row_string;
      if (!value) {
        return;
      }
      cell_value = {};
      if (_is_number(value)) {
        cell_value = {
          '$': {
            r: cell_name
          },
          v: [value]
        };
      } else {
        next_index = existing_setting && existing_setting[value] ? existing_setting[value] : this.shared_strings.add_string(value);
        cell_value = {
          '$': {
            r: cell_name,
            t: 's'
          },
          v: [next_index]
        };
      }
      row_string = _get_row_string(cell_name);
      row = this.row_by_name(sheetname, row_string);
      cell = this.cell_by_name(sheetname, cell_name);
      if (row === void 0) {
        new_row = {
          '$': {
            r: row_string,
            spans: '1:5'
          },
          c: [cell_value]
        };
        return sheet_xml.value.worksheet.sheetData[0].row.push(new_row);
      } else if (cell === void 0) {
        row.c.push(cell_value);
        return this._update_row(sheet_xml, row);
      } else {
        cell_value['$'].s = cell['$'].s;
        if (cell_value['$'].t) {
          cell['$'].t = 's';
          cell.v = [next_index];
        } else {
          cell.v = [value];
          if (cell['$'].t) {
            delete cell['$'].t;
          }
        }
        return this._update_cell(sheet_xml, row, cell);
      }
    };

    SpreadSheet.prototype.row_by_name = function(sheetname, row_number) {
      var row, sheet_xml;
      sheet_xml = this.sheet_by_name(sheetname);
      return row = _.find(sheet_xml.value.worksheet.sheetData[0].row, function(e) {
        return e['$'].r === row_number;
      });
    };

    SpreadSheet.prototype._update_row = function(sheet, row) {
      row.c = _.sortBy(row.c, function(e) {
        return _revert_number(_col_string(e['$'].r));
      });
      return _.each(sheet.value.worksheet.sheetData[0].row, function(existing_row) {
        if (existing_row['$'].r === row['$'].r) {
          return existing_row = row;
        }
      });
    };

    SpreadSheet.prototype._update_cell = function(sheet, row, cell) {
      row.c = _.sortBy(row.c, function(e) {
        return _revert_number(_col_string(e['$'].r));
      });
      return _.each(sheet.value.worksheet.sheetData[0].row, function(existing_row) {
        if (existing_row['$'].r === row['$'].r) {
          return _.each(existing_row.c, function(existing_cell) {
            if (existing_cell['$'].r === cell['$'].r) {
              return existing_cell = cell;
            }
          });
        }
      });
    };

    SpreadSheet.prototype._parse_dir_in_excel = function(dir) {
      var file_xmls, files;
      files = this.zip.folder(dir).file(/.xml/);
      file_xmls = [];
      return files.reduce((function(_this) {
        return function(promise, file) {
          return promise.then(function(prior_file) {
            return Promise.resolve().then(function() {
              return parseString(_this.zip.file(file.name).asText());
            }).then(function(file_xml) {
              file_xml.name = file.name.split('/')[file.name.split('/').length - 1];
              file_xmls.push(file_xml);
              return file_xmls;
            });
          });
        };
      })(this), Promise.resolve());
    };

    shared_strings = (function() {
      function shared_strings(obj) {
        this.add_string = bind(this.add_string, this);
        this.obj = obj;
        this.count = parseInt(obj.sst.si.length) - parseInt(1);
      }

      shared_strings.prototype.get_obj = function() {
        return this.obj;
      };

      shared_strings.prototype.add_string = function(value) {
        var new_string;
        if (!value) {
          value = '';
        }
        new_string = {
          t: [value],
          phoneticPr: [
            {
              '$': {
                fontId: '1'
              }
            }
          ]
        };
        this.obj.sst.si.push(new_string);
        return this.count = parseInt(this.obj.sst.si.length) - parseInt(1);
      };

      return shared_strings;

    })();

    return SpreadSheet;

  })();

  _is_number = function(value) {
    if (typeof value !== 'number' && typeof value !== 'string') {
      return false;
    } else {
      return value === parseFloat(value) && isFinite(value);
    }
  };

  _convert = function(value) {
    return 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('')[value];
  };

  _convert_alphabet = function(value) {
    var alphabet, alphabet1, alphabet2, alphabet3, number1, number2, number3;
    number1 = Math.floor(value / (26 * 26));
    number2 = Math.floor((value - number1 * 26 * 26) / 26);
    number3 = value - (number1 * 26 * 26 + number2 * 26);
    alphabet1 = _convert(number1) === 'A' ? '' : _convert(number1 - 1);
    alphabet2 = alphabet1 === '' && _convert(number2) === 'A' ? '' : _convert(number2 - 1);
    alphabet3 = _convert(number3);
    return alphabet = alphabet1 + alphabet2 + alphabet3;
  };

  _revert = function(alphabet) {
    return alphabet.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
  };

  _revert_number = function(alphabet) {
    var alphabet_with_zero, value;
    alphabet_with_zero = ('00' + alphabet).slice(-3).split('');
    return value = alphabet_with_zero[0] !== '0' ? value = value + _revert(alphabet_with_zero[0]) * 26 * 26 : alphabet_with_zero[1] !== '0' ? value = value + _revert(alphabet_with_zero[1]) * 26 : alphabet_with_zero[2] !== '0' ? value = value + _revert(alphabet_with_zero[2]) : 0;
  };

  _col_string = function(cell_name) {
    var cell_name_array, col_string, index;
    cell_name_array = cell_name.split('');
    index = 0;
    _.each(cell_name_array, function(c) {
      if (/^[a-zA-Z()]+$/.test(c)) {
        return index++;
      }
    });
    return col_string = cell_name.substr(0, index);
  };

  _get_row_string = function(cell_name) {
    var cell_name_array, index, row_string;
    cell_name_array = cell_name.split('');
    index = 0;
    _.each(cell_name_array, function(c) {
      if (/^[a-zA-Z()]+$/.test(c)) {
        return index++;
      }
    });
    return row_string = cell_name.substr(index, cell_name.length - index);
  };

  _get_col_string = function(cell_name) {
    var cell_name_array, col_string, index;
    cell_name_array = cell_name.split('');
    index = 0;
    _.each(cell_name_array, function(c) {
      if (/^[a-zA-Z()]+$/.test(c)) {
        return index++;
      }
    });
    return col_string = cell_name.substr(0, index);
  };

  load_config = function() {
    var config;
    return config = yaml.safeLoad(fs.readFileSync('./yaml/config.yml', 'utf8'));
  };

  module.exports = SpreadSheet;

}).call(this);
