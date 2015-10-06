# js-merge-xlsx
Minimum JavasScript-based template engine for MS-Excel.  

# Overview

js-merge-xlsx allows you to bind JavaScript values as follows.  

- Both on web browser and Node.js
- Bulk insert. Both of multi-file and multi-sheet is available.

Template  
![Template](https://raw.githubusercontent.com/hagasatoshi/js-merge-xlsx/master/image/before2.png)  
After rendering  
![Rendered](https://raw.githubusercontent.com/hagasatoshi/js-merge-xlsx/master/image/after.png)  

# Install

```
npm install js-merge-xlsx
```

# Usage(Node.js)

Prepare the template with bind-variables as mustache format {{}}.
![Template](https://raw.githubusercontent.com/hagasatoshi/js-merge-xlsx/master/image/before2.png)  

js-merge-xlsx supports Promises A+(bluebird). So, it is called basically in Promise-chain.  
Load your template using ExcelMerge#load() as follows.

```
fs.readFileAsync('./template/Template.xlsx')
.then((excel_template)=>{
    return Promise.props({
        rendering_data: readYamlAsync('./data/data.yml'),
        merge: new ExcelMerge().load(new JSZip(excel_template))
    });
```

Bind your data using ExcelMerge#render() as follows.
```
}).then((result)=>{
    let rendering_data = result.rendering_data;
    let merge =  result.merge;
    return merge.render(rendering_data);
}).then((excel_data)=>{
```

Please check [example codes](https://github.com/hagasatoshi/js-merge-xlsx/tree/master/example/1_node) and API below for the detail.

# Usage(on web browser)

You can also use it on web browser by using webpack(browserify). 
Bluebird automatically casts thenable objects, such as object returned by "$http.get()" or "$.get()", to trusted Promise instance. https://github.com/petkaantonov/bluebird/blob/master/API.md#promiseresolvedynamic-value---promise  
So you can use it in Promise-chain as well as on Node.js.  

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
            saveAs(excel_data,'Example.xlsx');
        }).catch((err)=>{
            console.error(err);
        });
```


