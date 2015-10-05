/**
 * * app.js
 * * Example for the usage of ExcelMerge#bulk_render_multi_sheet() on Node.js
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
import _ from 'underscore'

fsAsync.readFileAsync('./template/Template.xlsx')
.then((excel_template)=>{
    return Promise.props({
        rendering_data: readYamlAsync('./data/data.yml'),
        merge: new ExcelMerge().load(new JSZip(excel_template))
    });
}).then((result)=>{
    let rendering_data = [];
    _.each(result.rendering_data, (data,index)=>{
        rendering_data.push({name:'example'+(index+1)+'', data:data});
    });
    let merge =  result.merge;
    return merge.bulk_render_multi_sheet(rendering_data);
}).then((excel_data)=>{
    fsAsync.writeFileAsync('Example.xlsx',excel_data);
}).then(()=>{
    console.log('Success!!');
}).catch((err)=>{
    console.error(new Error(err).stack);
});