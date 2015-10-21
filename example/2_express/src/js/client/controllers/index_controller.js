/**
 * * indexController.js
 * * angular controller definition
 * * @author Satoshi Haga
 * * @date 2015/10/06
 **/

var Promise = require('bluebird');
var ExcelMerge = require('js-merge-xlsx');
var JSZip = require('jszip');
var _ = require('underscore');

var indexController = ($scope, $http)=>{

    /**
     * * exampleMerge
     * * example of ExcelMerge#merge()
     */
    $scope.exampleMerge = ()=>{
        Promise.resolve($http.get('/template/Template.xlsx', {responseType: "arraybuffer"}))
        .then((excelTemplate)=>{
            return Promise.props({
                data: $http.get('/data/data1.json'),
                excelMerge: new ExcelMerge().load(new JSZip(excelTemplate.data))
            });
        }).then(({data, excelMerge})=>{
            return excelMerge.merge(data.data);
        }).then((excelData)=>{
            saveAs(excelData,'example.xlsx');   //FileSaver#saveAs()
        }).catch((err)=>{
            console.error(err);
        });
    };

    /**
     * * exampleBulkMergeMultiFile
     * * example of ExcelMerge#bulkMergeMultiFile()
     */
    $scope.exampleBulkMergeMultiFile = ()=>{
        Promise.resolve($http.get('/template/Template.xlsx', {responseType: "arraybuffer"}))
        .then((excelTemplate)=>{
            return Promise.props({
                data: $http.get('/data/data2.json'),
                excelMerge: new ExcelMerge().load(new JSZip(excelTemplate.data))
            });
        }).then(({data, excelMerge})=>{
            data = _.map(data.data, (e,index)=>({name:`file${(index+1)}.xlsx`, data:e}));
            return excelMerge.bulkMergeMultiFile(data); //FileSaver#saveAs()
        }).then((zipData)=>{
            saveAs(zipData,'example.zip');
        }).catch((err)=>{
            console.error(err);
        });
    };

    /**
     * * exampleBulkMergeMultiSheet
     * * example of ExcelMerge#bulkMergeMultiSheet()
     */
    $scope.exampleBulkMergeMultiSheet = ()=>{
        Promise.resolve($http.get('/template/Template.xlsx', {responseType: "arraybuffer"}))
        .then((excelTemplate)=>{
            return Promise.props({
                data: $http.get('/data/data2.json'),
                excelMerge: new ExcelMerge().load(new JSZip(excelTemplate.data))
            });
        }).then(({data, excelMerge})=>{
            data = _.map(data.data, (e,index)=>({name:`sample${(index+1)}`, data:e}));
            return excelMerge.bulkMergeMultiSheet(data);
        }).then((excelData)=>{
            saveAs(excelData,'example.xlsx');
        }).catch((err)=>{
            console.error(err);
        });
    };

    $scope.charSingle = ()=>{
        Promise.resolve($http.get('/template/Template.xlsx', {responseType: "arraybuffer"}))
            .then((excelTemplate)=>{
                return Promise.props({
                    excelMerge: new ExcelMerge().load(new JSZip(excelTemplate.data))
                });
            }).then(({excelMerge})=>{
                let data = {
                    AccountName__c: `<>"'&'(0=0|~|==0~==0)=(('('&'%%&%%'$%$`,
                    AccountAddress__c: "KSPOI0)I0I0K0)(()')('#)JOKJ_?><<MNNBVCXXZ",
                    StartDateFormat__c: '2015/10/01',
                    EndDateFormat__c: '2016-9-30',
                    Address__c: '！イ”＝０｀M＝０｀イ＝『＝０『オ＝〜＝オ＝〜KW＝｀イ）＝｀０）！｜『？ァQ',
                    JobDescription__c: '①②③④⑤',
                    StartTime__c: '@@@@@@',
                    EndTime__c: '奧應橫歐毆穩假價畫會囘懷繪擴殼覺學嶽樂勸嵩'
                };
                return excelMerge.merge(data);
            }).then((excelData)=>{
                saveAs(excelData,'example.xlsx');   //FileSaver#saveAs()
            }).catch((err)=>{
                console.error(err);
            });
    };

    $scope.charBulkFiles = ()=>{
        Promise.resolve($http.get('/template/Template.xlsx', {responseType: "arraybuffer"}))
            .then((excelTemplate)=>{
                return Promise.props({
                    excelMerge: new ExcelMerge().load(new JSZip(excelTemplate.data))
                });
            }).then(({excelMerge})=>{
                let data = [
                    {
                        name:'file1.xlsx',
                        data:{
                            AccountName__c: `<>"'&'(0=0|~|==0~==0)=(('('&'%%&%%'$%$`,
                            AccountAddress__c: "KSPOI0)I0I0K0)(()')('#)JOKJ_?><<MNNBVCXXZ",
                            StartDateFormat__c: '2015/10/01',
                            EndDateFormat__c: '2016-9-30',
                            Address__c: '！イ”＝０｀M＝０｀イ＝『＝０『オ＝〜＝オ＝〜KW＝｀イ）＝｀０）！｜『？ァQ',
                            JobDescription__c: '①②③④⑤',
                            StartTime__c: '@@@@@@',
                            EndTime__c: '奧應橫歐毆穩假價畫會囘懷繪擴殼覺學嶽樂勸嵩'
                        }
                    },
                    {
                        name:'file2.xlsx',
                        data:{
                            AccountName__c: `<>"'&'(0=0|~|==0~==0)=(('('&'%%&%%'$%$`,
                            AccountAddress__c: "KSPOI0)I0I0K0)(()')('#)JOKJ_?><<MNNBVCXXZ",
                            StartDateFormat__c: '2015/10/01',
                            EndDateFormat__c: '2016-9-30',
                            Address__c: '！イ”＝０｀M＝０｀イ＝『＝０『オ＝〜＝オ＝〜KW＝｀イ）＝｀０）！｜『？ァQ',
                            JobDescription__c: '①②③④⑤',
                            StartTime__c: '@@@@@@',
                            EndTime__c: '奧應橫歐毆穩假價畫會囘懷繪擴殼覺學嶽樂勸嵩'
                        }
                    },
                    {
                        name:'file3.xlsx',
                        data:{
                            AccountName__c: `<>"'&'(0=0|~|==0~==0)=(('('&'%%&%%'$%$`,
                            AccountAddress__c: "KSPOI0)I0I0K0)(()')('#)JOKJ_?><<MNNBVCXXZ",
                            StartDateFormat__c: '2015/10/01',
                            EndDateFormat__c: '2016-9-30',
                            Address__c: '！イ”＝０｀M＝０｀イ＝『＝０『オ＝〜＝オ＝〜KW＝｀イ）＝｀０）！｜『？ァQ',
                            JobDescription__c: '①②③④⑤',
                            StartTime__c: '@@@@@@',
                            EndTime__c: '奧應橫歐毆穩假價畫會囘懷繪擴殼覺學嶽樂勸嵩'
                        }
                    }
                ];

                return excelMerge.bulkMergeMultiFile(data); //FileSaver#saveAs()
            }).then((zipData)=>{
                saveAs(zipData,'example.zip');
            }).catch((err)=>{
                console.error(err);
            });
    };

    $scope.charBulkSheets = ()=>{
        Promise.resolve($http.get('/template/Template.xlsx', {responseType: "arraybuffer"}))
            .then((excelTemplate)=>{
                return Promise.props({
                    excelMerge: new ExcelMerge().load(new JSZip(excelTemplate.data))
                });
            }).then(({excelMerge})=>{
                let data = [
                    {
                        name:'sheet1',
                        data:{
                            AccountName__c: `<>"'&'(0=0|~|==0~==0)=(('('&'%%&%%'$%$`,
                            AccountAddress__c: "KSPOI0)I0I0K0)(()')('#)JOKJ_?><<MNNBVCXXZ",
                            StartDateFormat__c: '2015/10/01',
                            EndDateFormat__c: '2016-9-30',
                            Address__c: '！イ”＝０｀M＝０｀イ＝『＝０『オ＝〜＝オ＝〜KW＝｀イ）＝｀０）！｜『？ァQ',
                            JobDescription__c: '①②③④⑤',
                            StartTime__c: '@@@@@@',
                            EndTime__c: '奧應橫歐毆穩假價畫會囘懷繪擴殼覺學嶽樂勸嵩'
                        }
                    },
                    {
                        name:'sheet2',
                        data:{
                            AccountName__c: `<>"'&'(0=0|~|==0~==0)=(('('&'%%&%%'$%$`,
                            AccountAddress__c: "KSPOI0)I0I0K0)(()')('#)JOKJ_?><<MNNBVCXXZ",
                            StartDateFormat__c: '2015/10/01',
                            EndDateFormat__c: '2016-9-30',
                            Address__c: '！イ”＝０｀M＝０｀イ＝『＝０『オ＝〜＝オ＝〜KW＝｀イ）＝｀０）！｜『？ァQ',
                            JobDescription__c: '①②③④⑤',
                            StartTime__c: '@@@@@@',
                            EndTime__c: '奧應橫歐毆穩假價畫會囘懷繪擴殼覺學嶽樂勸嵩'
                        }
                    },
                    {
                        name:'sheet3',
                        data:{
                            AccountName__c: `<>"'&'(0=0|~|==0~==0)=(('('&'%%&%%'$%$`,
                            AccountAddress__c: "KSPOI0)I0I0K0)(()')('#)JOKJ_?><<MNNBVCXXZ",
                            StartDateFormat__c: '2015/10/01',
                            EndDateFormat__c: '2016-9-30',
                            Address__c: '！イ”＝０｀M＝０｀イ＝『＝０『オ＝〜＝オ＝〜KW＝｀イ）＝｀０）！｜『？ァQ',
                            JobDescription__c: '①②③④⑤',
                            StartTime__c: '@@@@@@',
                            EndTime__c: '奧應橫歐毆穩假價畫會囘懷繪擴殼覺學嶽樂勸嵩'
                        }
                    }
                ];
                return excelMerge.bulkMergeMultiSheet(data);
            }).then((excelData)=>{
                saveAs(excelData,'example.xlsx');
            }).catch((err)=>{
                console.error(err);
            });
    };
};

module.exports = indexController;