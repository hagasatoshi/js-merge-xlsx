# js-merge-xlsx  
Minimum JavasScript-based template engine for MS-Excel. js-merge-xlsx allows you to print JavaScript values.  

- Avairable for both web browser and Node.js .
- Bulk printing. It is possible to print array as 'multiple files'. 
- Bulk printing. It is possible to print array as 'multiple sheets'. 

Template  
![Template](https://raw.githubusercontent.com/hagasatoshi/js-merge-xlsx/master/image/before2.png)  
After printing  
![Rendered](https://raw.githubusercontent.com/hagasatoshi/js-merge-xlsx/master/image/after.png)  

# Install
```
npm install js-merge-xlsx
```

# Prepare template  
Prepare the template with bind-variables as mustache format {{}}.
![Template](https://raw.githubusercontent.com/hagasatoshi/js-merge-xlsx/master/image/before2.png)  

# Usage(Node.js)  
js-merge-xlsx supports Promises A+(bluebird). So, it is called basically in Promise-chain.  
ExcelMerge#load() ES6 syntax  
```
fs.readFileAsync('./template/Template.xlsx')
.then((excel_template)=>{
    return Promise.props({
        rendering_data: readYamlAsync('./data/data.yml'),
        merge: new ExcelMerge().load(new JSZip(excel_template))
    });
```

ExcelMerge#render() ES6 syntax  
```
}).then((result)=>{
    let rendering_data = result.rendering_data;
    let merge =  result.merge;
    return merge.render(rendering_data);
}).then((excel_data)=>{
```

Please check [example codes](https://github.com/hagasatoshi/js-merge-xlsx/tree/master/example/1_node) and API below for detail.

# Usage(on web browser)
You can also use it on web browser by using webpack(browserify). 
Bluebird automatically casts thenable object, such as object returned by "$http.get()" or "$.get()", to trusted Promise. https://github.com/petkaantonov/bluebird/blob/master/API.md#promiseresolvedynamic-value---promise  
So you can use it in Promise-chain as well as on Node.js.  
Example(ES6 syntax)  
```
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
    saveAs(excel_data,'Example.xlsx');  //FileSaver
}).catch((err)=>{
    console.error(err);
});
```

Please check [example codes](https://github.com/hagasatoshi/js-merge-xlsx/tree/master/example/2_express) and API below for detail.