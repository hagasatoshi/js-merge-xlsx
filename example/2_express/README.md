# Example on web browser  
Here is example on web browser. Webpack(or browserify) empower you to use node-modules on web-browser as well. Also bluebird automatically casts thenable object, such as object returned by '$http.get()' and '$.get()', to trusted Promise. So, you can code in the same way as Node.js.  
  
# git clone
```bash
git clone git@github.com:hagasatoshi/js-merge-xlsx.git
cd js-merge-xlsx/example/2_express
```
# Build
```bash
npm install
bower install
gulp
```
# Start server
```bash
node server.js
```

You can access 'http://localhost:3000/' as follows.  
![Experss](https://raw.githubusercontent.com/hagasatoshi/js-merge-xlsx/master/image/express.png)  
  
Each button prints using js-merge-xlsx.

- ExcelMerge#render() : single printing.
- ExcelMerge#bulk_render_multi_file() : bulk printing as 'multiple file'.
- ExcelMerge#bulk_render_multi_sheet() : bulk printing as 'multiple sheet'.
  
# Source  
Angular's controller(ES6 syntax)
```JavaScript
/**
 * * index_controller.js
 * * angular controller definition
 * * @author Satoshi Haga
 * * @date 2015/10/06
 **/

import Promise from 'bluebird'
import ExcelMerge from 'js-merge-xlsx'
import JSZip from 'jszip'
import _ from 'underscore'

var index_controller = ($scope, $http)=>{

    /**
     * * example_render
     * * example of ExcelMerge#render()
     */
    $scope.example_render = ()=>{
        Promise.resolve($http.get('/template/Template.xlsx', {responseType: "arraybuffer"}))
        .then((excel_template)=>{
            return Promise.props({
                rendering_data: $http.get('/data/data1.json'),
                merge: new ExcelMerge().load(new JSZip(excel_template.data))
            });
        }).then((result)=>{
            let rendering_data = result.rendering_data.data;
            let merge =  result.merge;
            return merge.render(rendering_data);
        }).then((excel_data)=>{
            saveAs(excel_data,'Example.xlsx');
        }).catch((err)=>{
            console.error(err);
        });
    };

    /**
     * * example_bulk_render_multi_file
     * * example of ExcelMerge#bulk_render_multi_file()
     */
    $scope.example_bulk_render_multi_file = ()=>{
        Promise.resolve($http.get('/template/Template.xlsx', {responseType: "arraybuffer"}))
            .then((excel_template)=>{
                return Promise.props({
                    rendering_data: $http.get('/data/data2.json'),
                    merge: new ExcelMerge().load(new JSZip(excel_template.data))
                });
            }).then((result)=>{
                let rendering_data = [];
                _.each(result.rendering_data.data, (data,index)=>{
                    rendering_data.push({name:'file'+(index+1)+'.xlsx', data:data});
                });
                let merge =  result.merge;
                return merge.bulk_render_multi_file(rendering_data);
            }).then((zip_data)=>{
                saveAs(zip_data,'Example.zip');
            }).catch((err)=>{
                console.error(err);
            });
    };

    /**
     * * example_bulk_render_multi_sheet
     * * example of ExcelMerge#bulk_render_multi_sheet()
     */
    $scope.example_bulk_render_multi_sheet = ()=>{
        Promise.resolve($http.get('/template/Template.xlsx', {responseType: "arraybuffer"}))
            .then((excel_template)=>{
                return Promise.props({
                    rendering_data: $http.get('/data/data2.json'),
                    merge: new ExcelMerge().load(new JSZip(excel_template.data))
                });
            }).then((result)=>{
                let rendering_data = [];
                _.each(result.rendering_data.data, (data,index)=>{
                    rendering_data.push({name:'sample'+(index+1), data:data});
                });
                let merge =  result.merge;
                return merge.bulk_render_multi_sheet(rendering_data);
            }).then((excel_data)=>{
                saveAs(excel_data,'Example.xlsx');
            }).catch((err)=>{
                console.error(err);
            });
    };
};

module.exports = index_controller;
```
