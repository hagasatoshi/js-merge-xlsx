const Promise = require('bluebird');
const {merge, bulkMergeToFiles, bulkMergeToSheets} = require('js-merge-xlsx');
const JSZip = require('jszip');
const _ = require('underscore');

module.exports = ($scope, $http) => {

    $scope.merge = () => {
        Promise.props({
            template: $http.get('/template/Template.xlsx', {responseType: 'arraybuffer'}),
            data:     $http.get('/data/data1.json')
        }).then(({template, data}) => {

            //FileSaver#saveAs()
            saveAs(merge(template, data), 'example.xlsx');
        }).catch((err) => {
            console.log(err);
        });
    };

    $scope.bulkMergeToFiles = () => {
        Promise.props({
            template: $http.get('/template/Template.xlsx', {responseType: 'arraybuffer'}),
            data:     $http.get('/data/data2.json')
        }).then(({template, data}) => {

            data = _.map(data.data, (e,index) => {
                return {name: `file${(index+1)}.xlsx`, data: e};
            });
            //FileSaver#saveAs()
            saveAs(bulkMergeToFiles(template, data), 'example.zip');
        }).catch((err) => {
            console.log(err);
        });
    };

    $scope.bulkMergeToSheets = ()=>{
        Promise.props({
            template: $http.get('/template/Template.xlsx', {responseType: 'arraybuffer'}),
            data:     $http.get('/data/data2.json')
        }).then(({template, data}) => {

            data = _.map(data.data, (e,index) => {
                return {name: `sample${(index+1)}`, data: e};
            });

            //bulkMergeToSheets() is called asyc by returning Promise(bluebird) instance.
            return bulkMergeToSheets(template, data);
        }).then((excel) => {

            saveAs(excel,'example.xlsx');
        }).catch((err) => {
            console.log(err);
        });
    };
};