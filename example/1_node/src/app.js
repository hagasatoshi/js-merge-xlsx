/**
 * * app.js
 * * Example for the usage on Node.js
 * * @author Satoshi Haga
 * * @date 2015/09/30
 **/

import ExcelMerge from 'js-merge-xlsx'
import fs from 'fs'
import JSZip from 'jszip'
import _ from 'underscore'
import 'colors'


//Init template engine instance
var excel_data = fs.readFileSync('./Template.xlsx');
new ExcelMerge().load(new JSZip(excel_data))
    .then((merge)=> {
        return merge.bulk_render_multi_sheet([
            {
                name:'sheet11',
                data:{
                    AccountName__c: 'example1 corporation',
                    AccountAddress__c: 'US',
                    StartDateFormat__c: '2015/01/01'
                }
            },
            {
                name:'sheet12',
                data:{
                    AccountName__c: 'example2 corporation',
                    AccountAddress__c: 'US',
                    StartDateFormat__c: '2015/01/01'
                }
            },
            {
                name:'sheet13',
                data:{
                    AccountName__c: 'example3 corporation',
                    AccountAddress__c: 'US',
                    StartDateFormat__c: '2015/01/01'
                }
            }
        ],{type: "nodebuffer",compression:"DEFLATE"});
    }).then((excel_data)=>{
        fs.writeFileSync('sample.xlsx',excel_data);
    }).catch((err)=>{
        console.error(new Error(err).stack.red);
    });