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