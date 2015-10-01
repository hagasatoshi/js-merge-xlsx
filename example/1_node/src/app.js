/**
 * * app.js
 * * Example for the usage on Node.js
 * * @author Satoshi Haga
 * * @date 2015/09/30
 **/

import ExcelMerge from 'js-merge-xlsx'
import fs from 'fs'
import JSZip from 'jszip'

//Init template engine instance
var excel_data = fs.readFileSync('./template/Template.xlsx');
var merge = new ExcelMerge(new JSZip(excel_data));

//Prepare binding-data
var example_data = {
    AccountName__c: 'example corporation',
    AccountAddress__c: 'US',
    StartDateFormat__c: '2015/01/01'
};

//Bind data
var rendered_data = merge.render(example_data,{type: "nodebuffer",compression:"DEFLATE"});
fs.writeFileSync('./RederedData.xlsx',rendered_data);
