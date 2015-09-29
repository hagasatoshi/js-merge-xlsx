/**
 * * ExcelMerge
 * * Template managing class. wrapping JsZip object.
 * * @author Satoshi Haga
 * * @date 2015/09/30
 **/

"use strict";

var _createClass = (function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; })();

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { "default": obj }; }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _mustache = require('mustache');

var _mustache2 = _interopRequireDefault(_mustache);

var ExcelMerge = (function () {

  /**
   * * constructor
   * * @param {Object} excel JsZip object including MS-Excel file
   **/

  function ExcelMerge(excel) {
    _classCallCheck(this, ExcelMerge);

    this.excel = excel;
  }

  //Exports

  /**
   * * render
   * * @param {Object} bind_data binding data
   * * @param {Object} jszip_option JsZip#generate() option.
   * * @returns {Object} rendered MS-Excel data. data-format is determined by jszip_option
   **/

  _createClass(ExcelMerge, [{
    key: "render",
    value: function render(bind_data) {
      var jszip_option = arguments.length <= 1 || arguments[1] === undefined ? { type: "blob", compression: "DEFLATE" } : arguments[1];

      var template = this.excel.file('xl/sharedStrings.xml').asText();
      this.excel.file('xl/sharedStrings.xml', _mustache2["default"].render(template, bind_data));
      return this.excel.generate(jszip_option);
    }
  }]);

  return ExcelMerge;
})();

module.exports = ExcelMerge;