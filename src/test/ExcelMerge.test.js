const Promise = require('bluebird');
const _ = require('underscore');
const readYamlAsync = Promise.promisify(require('read-yaml'));
const fs = Promise.promisifyAll(require('fs'));
const excelmerge = require('../ExcelMerge');
const assert = require('assert');

const config = {
    templateDir: './test/templates/',
    testDataDir: './test/data/',
    outptutDir:  './test/output/'
};

const removeExstingFiles = (done) => {
    _.each(fs.readdirSync(config.outptutDir), (file) => {
        fs.unlinkSync(`${config.outptutDir}${file}`);
    });
    done && done();
};

const readFiles = (template, yaml) => {
    return Promise.props({
        template: fs.readFileAsync(`${config.templateDir}${template}`),
        data:     readYamlAsync(`${config.testDataDir}${yaml}`)
    });
};

removeExstingFiles();

describe('test for excelmerge.merge()', () => {

    it('excel file is created successfully', () => {
        return readFiles('Template.xlsx', 'data1.yml')
            .then(({template, data}) => {
                return fs.writeFileAsync(
                    `${config.outptutDir}test1.xlsx`, excelmerge.merge(template, data)
                );
            }).then(() => {
                assert(true);
            }).catch((err) => {
                console.log(err);
                assert(false);
            });
    });
});

describe('test for excelmerge.bulkMergeToFiles()', () => {

    it('excel file is created successfully', () => {
        return readFiles('Template.xlsx', 'data2.yml')
            .then(({template, data}) => {
                let arrayObj = _.map(
                    data, (e, index) => ({name: `file${index}.xlsx`, data: e})
                );
                return fs.writeFileAsync(
                    `${config.outptutDir}test2.zip`,
                    excelmerge.bulkMergeToFiles(template, arrayObj)
                );
            }).then(() => {
                assert(true);
            }).catch((err) => {
                console.log(err);
                assert(false);
            });
    });
});

describe('test for excelmerge.bulkMergeToSheets()', () => {

    it('excel file is created successfully', () => {
        return readFiles('Template.xlsx', 'data2.yml')
            .then(({template, data}) => {
                let arrayObj = _.map(
                    data, (e, index) => ({name: `test${index}`, data: e})
                );
                return excelmerge.bulkMergeToSheets(template, arrayObj);
            }).then((excelData) => {
                return fs.writeFileAsync(
                    `${config.outptutDir}test3.xlsx`, excelData
                );
            }).then(() => {
                assert(true);
            }).catch((err) => {
                console.log(err);
                assert(false);
            });
    });
});