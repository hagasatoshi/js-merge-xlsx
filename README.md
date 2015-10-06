# js-merge-xlsx  
Minimum JavaScript-based template engine for MS-Excel. js-merge-xlsx empowers you to print JavaScript objects.

- Available for both web browser and Node.js .
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
Note: Only string cell is supported. Please make sure that the format of cells having variables is STRING.  
![Note](https://raw.githubusercontent.com/hagasatoshi/js-merge-xlsx/master/image/cell_format.png)

# Node.js  
js-merge-xlsx supports Promises/A+([bluebird](https://github.com/petkaantonov/bluebird)). So, it is called basically in Promise-chain.  
Example(ES6 syntax)  
```
fs.readFileAsync('./template/Template.xlsx')
.then((excel_template)=>{
    return Promise.props({
        rendering_data: readYamlAsync('./data/data.yml'),
        merge: new ExcelMerge().load(new JSZip(excel_template))
    });
}).then((result)=>{
    let rendering_data = result.rendering_data;
    let merge =  result.merge;
    return merge.render(rendering_data);
}).then((excel_data)=>{
    fs.writeFileAsync('Example.xlsx',excel_data);
}).catch((err)=>{
    console.error(new Error(err).stack);
});
```

Please check [example codes](https://github.com/hagasatoshi/js-merge-xlsx/tree/master/example/1_node) and API below for detail.

# Browser  
You can also use it on web browser by using webpack(browserify). 
Bluebird automatically casts thenable object, such as object returned by "$http.get()" or "$.get()", to trusted Promise. https://github.com/petkaantonov/bluebird/blob/master/API.md#promiseresolvedynamic-value---promise  
So, you can code in the same way as Node.js.    
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
