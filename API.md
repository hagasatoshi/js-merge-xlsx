# API Reference  
  
- [Initialize](#core)
    - [`new ExcelMerge()`](#new-promisefunctionfunction-resolve-function-reject-resolver---promise)
    - [`.load(JSZip zip)`](#thenfunction-fulfilledhandler--function-rejectedhandler----promise)
- [Rendering](#core)
    - [`.render()`](#new-promisefunctionfunction-resolve-function-reject-resolver---promise)
    - [`.bulk_render_multi_file()`](#new-promisefunctionfunction-resolve-function-reject-resolver---promise)
    - [`.bulk_render_multi_sheet()`](#new-promisefunctionfunction-resolve-function-reject-resolver---promise)
  
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

#####`bulk_render_multi_file([{name:name of file1, data:data of file1},{name:name of file1, data:data of file1},,,])` -> `Promise`  
Render array as multiple files. Returns Promise instance including Zip-file data. If on Node.js, the type of data is Buffer instance. If on web browser, blob is returned.  
```
```
