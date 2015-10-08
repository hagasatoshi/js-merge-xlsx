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
.then((excel_template)=>{
    return new ExcelMerge().load(new JSZip(excel_template)); //Initialize ExcelMerge object
}).then((merge)=>{
    //merge is ExcelMerge instance.
```
# Render    
#####`render(Object data)` -> `Promise`  
Render single object, not array. Returns Promise instance including MS-Excel data. If on Node.js, the type of data is Buffer instance. If on web browser, blob is returned.
```
    return merge.render(some_data);
}).then((excel_data)=>{
    fs.writeFileAsync('Example1.xlsx',excel_data);
}).catch((err)=>{
    console.error(err);
});
```

#####`bulk_render_multi_file([{name:file's name of file1, data:data of file1},{name:file's name of file2, data:data of file2},,,])` -> `Promise`  
Render array as multiple files. Returns Promise instance including Zip-file data. If on Node.js, the type of data is Buffer instance. If on web browser, blob is returned. You can use 'name' property as each file-name in zip file.
```
  let array_hoge = [];
  _.each(array_fuga, (data,index)=>{
      array_hoge.push({name:'example'+(index+1), data:data});
  });
  return merge.bulk_render_multi_file(array_hoge);
}).then((zip_data)=>{
  fs.writeFileAsync('piyo.zip',zip_data);
}).catch((err)=>{
    console.error(err);
});
```

#####`bulk_render_multi_sheet([{name:sheet's name of file1, data:data of file1},{name:sheet's name of file1, data:data of file1},,,])` -> `Promise`  
Render array as multiple sheet. Returns Promise instance including Excel-file data. If on Node.js, the type of data is Buffer instance. If on web browser, blob is returned. You can use 'name' property as each sheet-name in MS-Excel file.
```
  let array_hoge = [];
  _.each(array_fuga, (data,index)=>{
      array_hoge.push({name:'example'+(index+1), data:data});
  });
  return merge.bulk_render_multi_sheet(array_hoge);
}).then((excel_data)=>{
  fs.writeFileAsync('piyo.xlsx',excel_data);
}).catch((err)=>{
    console.error(err);
});
```
