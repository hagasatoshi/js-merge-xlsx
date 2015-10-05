/**
 * * app.js
 * * Example for the usage of ExcelMerge#render() on Node.js
 * * @author Satoshi Haga
 * * @date 2015/09/30
 **/

'use strict';

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { 'default': obj }; }

var _jsMergeXlsx = require('js-merge-xlsx');

var _jsMergeXlsx2 = _interopRequireDefault(_jsMergeXlsx);

var _bluebird = require('bluebird');

var _bluebird2 = _interopRequireDefault(_bluebird);

var _readYaml = require('read-yaml');

var _readYaml2 = _interopRequireDefault(_readYaml);

var _fs = require('fs');

var _fs2 = _interopRequireDefault(_fs);

var _jszip = require('jszip');

var _jszip2 = _interopRequireDefault(_jszip);

var readYamlAsync = _bluebird2['default'].promisify(_readYaml2['default']);

var fsAsync = _bluebird2['default'].promisifyAll(_fs2['default']);

fsAsync.readFileAsync('./template/Template.xlsx').then(function (excel_template) {
    return _bluebird2['default'].props({
        rendering_data: readYamlAsync('./data/data.yml'),
        merge: new _jsMergeXlsx2['default']().load(new _jszip2['default'](excel_template))
    });
}).then(function (result) {
    var rendering_data = result.rendering_data;
    var merge = result.merge;
    return merge.render(rendering_data);
}).then(function (excel_data) {
    fsAsync.writeFileAsync('Example.xlsx', excel_data);
}).then(function () {
    console.log('Success!!');
})['catch'](function (err) {
    console.error(new Error(err).stack);
});