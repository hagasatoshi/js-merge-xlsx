const Promise = require('bluebird');
const readYamlSync = require('read-data').yaml.sync;
const fs = Promise.promisifyAll(require('fs'));
const _ = require('underscore');
const {merge, bulkMergeToFiles, bulkMergeToSheets}
    = require('js-merge-xlsx');

const config = {
    template:   './template/Template.xlsx',
    singleData: './data/data1.yml',
    arrayData:  './data/data2.yml'
};

const readData = () => {
    let templateObj = fs.readFileSync(config.template);
    let data  = readYamlSync(config.singleData);
    let bulkData = readYamlSync(config.arrayData);

    return {
        templateObj: templateObj,
        data: data,
        bulkData1: _.map(bulkData, (e, index) => {
            return {name: `file${index + 1}.xlsx`, data: e};
        }),
        bulkData2: _.map(bulkData, (e, index) => {
            return {name: `example${index + 1}`, data: e};
        })
    };
};

//Start
let {templateObj, data, bulkData1, bulkData2} = readData();

//example of merge()
fs.writeFileSync('example1.xlsx',  merge(templateObj, data));

//example of bulkMergeToFiles()
fs.writeFileSync(
    'example2.zip',
    bulkMergeToFiles(templateObj, bulkData1)
);

//example of bulkMergeToSheets()
//this method is called async by returning Promise(bluebird) instance.
bulkMergeToSheets(templateObj, bulkData2)
.then((excel) => {
    fs.writeFileSync('example3.xlsx', excel);
});
