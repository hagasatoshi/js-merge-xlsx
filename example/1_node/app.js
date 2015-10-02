/**
 * * app.js
 * * Example for the usage on Node.js
 * * @author Satoshi Haga
 * * @date 2015/09/30
 **/

'use strict';

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { 'default': obj }; }

var _jsMergeXlsx = require('js-merge-xlsx');

var _jsMergeXlsx2 = _interopRequireDefault(_jsMergeXlsx);

var _fs = require('fs');

var _fs2 = _interopRequireDefault(_fs);

var _jszip = require('jszip');

var _jszip2 = _interopRequireDefault(_jszip);

//Init template engine instance
var excel_data = _fs2['default'].readFileSync('./template/Template.xlsx');
var merge = new _jsMergeXlsx2['default'](new _jszip2['default'](excel_data));

//Prepare binding-data
var example_data = {
  AccountName__c: 'example corporation',
  AccountAddress__c: 'US',
  StartDateFormat__c: '2015/01/01'
};

//Bind data
var rendered_data = merge.render(example_data, { type: "nodebuffer", compression: "DEFLATE" });
_fs2['default'].writeFileSync('./RederedData.xlsx', rendered_data);