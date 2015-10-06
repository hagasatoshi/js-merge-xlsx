/**
 * * app.js
 * * Example on Node.js
 * * @author Satoshi Haga
 * * @date 2015/09/30
 **/

'use strict';

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { 'default': obj }; }

var _jsMergeXlsx = require('js-merge-xlsx');

var _jsMergeXlsx2 = _interopRequireDefault(_jsMergeXlsx);

var _bluebird = require('bluebird');

var _bluebird2 = _interopRequireDefault(_bluebird);

var _jszip = require('jszip');

var _jszip2 = _interopRequireDefault(_jszip);

var _underscore = require('underscore');

var _underscore2 = _interopRequireDefault(_underscore);

//Load Template

var readYamlAsync = _bluebird2['default'].promisify(require('read-yaml'));
var fs = _bluebird2['default'].promisifyAll(require('fs'));
fs.readFileAsync('./template/Template.xlsx').then(function (excel_template) {
    return _bluebird2['default'].props({
        rendering_data1: readYamlAsync('./data/data1.yml'), //Load single data
        rendering_data2: readYamlAsync('./data/data2.yml'), //Load array data
        merge: new _jsMergeXlsx2['default']().load(new _jszip2['default'](excel_template)) //Initialize ExcelMerge object
    });
}).then(function (result) {
    //Single-printing
    var rendering_data1 = result.rendering_data1;

    //Bulk-printing as 'multiple files'
    var rendering_data2 = [];
    _underscore2['default'].each(result.rendering_data2, function (data, index) {
        rendering_data2.push({ name: 'file' + (index + 1) + '.xlsx', data: data });
    });

    //Bulk-printing as 'multiple sheets'
    var rendering_data3 = [];
    _underscore2['default'].each(result.rendering_data2, function (data, index) {
        rendering_data3.push({ name: 'example' + (index + 1), data: data });
    });

    //ExcelMerge object
    var merge = result.merge;

    //Execute rendering
    return _bluebird2['default'].props({
        excel_data1: merge.render(rendering_data1),
        excel_data2: merge.bulk_render_multi_file(rendering_data2),
        excel_data3: merge.bulk_render_multi_sheet(rendering_data3)
    });
}).then(function (result) {
    return _bluebird2['default'].all([fs.writeFileAsync('Example1.xlsx', result.excel_data1), fs.writeFileAsync('Example2.zip', result.excel_data2), fs.writeFileAsync('Example3.xlsx', result.excel_data3)]);
})['catch'](function (err) {
    console.error(new Error(err).stack);
});