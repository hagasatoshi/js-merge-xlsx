/**
 * * app.js
 * * Example for the usage of ExcelMerge#render()  on Node.js
 * * @author Satoshi Haga
 * * @date 2015/09/30
 **/

import ExcelMerge from 'js-merge-xlsx'
import Promise from 'bluebird'
import readYaml from 'read-yaml'
var readYamlAsync = Promise.promisify(readYaml);
import fs from 'fs'
var fsAsync = Promise.promisifyAll(fs);
import JSZip from 'jszip'
import 'colors'

fsAsync.readFileAsync('./template/Template.xlsx')
.then((excel_template)=>{
    return Promise.props({
        rendering_data: readYamlAsync('./data/data.yml'),
        merge: new ExcelMerge().load(new JSZip(excel_template))
    });
}).then((result)=>{
    let rendering_data = result.rendering_data;
    let merge =  result.merge;
    return merge.render(rendering_data, {type: "nodebuffer",compression:"DEFLATE"});
}).then((excel_data)=>{
    fsAsync.writeFileAsync('Example.xlsx',excel_data);
}).then(()=>{
    console.log('Success!!');
}).catch((err)=>{
    console.error(new Error(err).stack.red);
});