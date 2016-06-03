/**
 * Config
 * @author Satoshi Haga
 * @date 2016/03/27
 */

'use strict';

var isNode = require('detect-node');

module.exports = {
    EXCEL_FILES: {
        FILE_SHARED_STRINGS: 'xl/sharedStrings.xml',
        FILE_WORKBOOK_RELS: 'xl/_rels/workbook.xml.rels',
        FILE_WORKBOOK: 'xl/workbook.xml',
        DIR_WORKSHEETS: 'xl/worksheets',
        DIR_WORKSHEETS_RELS: 'xl/worksheets/_rels'
    },
    JSZIP_OPTION: {
        COMPLESSION: 'DEFLATE',
        BUFFER_TYPE_OUTPUT: isNode ? 'nodebuffer' : 'blob',
        BUFFER_TYPE_JSZIP: isNode ? 'nodebuffer' : 'arraybuffer'
    },
    TEST_DIRS: {
        TEMPLATE: './test/templates/',
        DATA: './test/data/',
        OUTPUT: './test/output/',
        XML: './test/xml/'
    },
    OPEN_XML_SCHEMA_DEFINITION: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
};