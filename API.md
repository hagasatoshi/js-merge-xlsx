# API Reference  
  
- [Initialize](#initialize)
    - [`new ExcelMerge()`](#new-excelmerge---excelmerge)
    - [`.load(JSZip zip)`](#loadjszip-zip---promise)
- [Render](#render)
    - [`.render()`](#renderobject-data---promise)
    - [`.bulk_render_multi_file()`](#bulk_render_multi_filenamefiles-name-of-file1-datadata-of-file1namefiles-name-of-file2-datadata-of-file2---promise)
    - [`.bulk_render_multi_sheet()`](#bulk_render_multi_sheetnamesheets-name-of-file1-datadata-of-file1namesheets-name-of-file1-datadata-of-file1---promise)

Each code example is written with ES6 syntax
# Initialize  
Create ExcelMerge instance and load Excel data.
#####`new ExcelMerge()` -> `ExcelMerge`  
Contructor. No arguments are required. 

#####`load(JSZip zip)` -> `Promise`  
Load MS-Excel template. Parameter is JSZip instance including MS-Excel data. Returns a new promise instance including this ExcelMerge instance. So you can code like method-chain as follows.  
```
fs.readFileAsync('./Template.xlsx')
.then((excelTemplate)=>{
    return new ExcelMerge().load(new JSZip(excelTemplate)); 
}).then((excelMerge)=>{
    //excelMerge is ExcelMerge instance.
```
# Render    
#####`merge(Object data)` -> `Promise`  
Render single object, not array. Returns Promise instance including MS-Excel data. If on Node.js, the type of data is Buffer instance. If on web browser, blob is returned.
```
  return excelMerge.merge(someData);
}).then((excelData)=>{
  fs.writeFileAsync('example1.xlsx',excelData);
}).catch((err)=>{
  console.error(err);
});
```

#####`bulkMergeMultiFile([{name:file's name of file1, data:data of file1},{name:file's name of file2, data:data of file2},,,])` -> `Promise`  
Render array as multiple files. Returns Promise instance including Zip-file data. If on Node.js, the type of data is Buffer instance. If on web browser, blob is returned. You can use 'name' property as each file-name in zip file.
```
  let arrayHoge = _.map(arrayFuga, (data,index)=>({name:`example${(index+1)}.xlsx`, data:data}));
  return excelMerge.bulkRenderMultiFile(arrayHoge);
}).then((zipData)=>{
  fs.writeFileAsync('piyo.zip',zipData);
}).catch((err)=>{
  console.error(err);
});
```

#####`bulkRenderMultiSheet([{name:sheet's name of file1, data:data of file1},{name:sheet's name of file1, data:data of file1},,,])` -> `Promise`  
Render array as multiple sheet. Returns Promise instance including Excel-file data. If on Node.js, the type of data is Buffer instance. If on web browser, blob is returned. You can use 'name' property as each sheet-name in MS-Excel file.
```
  let arrayHoge = _.map(arrayFuga, (data,index)=>({name:`example${(index+1)}`, data:data}));
  return merge.bulkRenderMultiSheet(arrayHoge);
}).then((excelData)=>{
  fs.writeFileAsync('piyo.xlsx',excelData);
}).catch((err)=>{
  console.error(err);
});
```
