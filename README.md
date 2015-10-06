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

# Usage

js-merge-xlsx supports Promises A+(bluebird). So, it is called in Promise-chain.

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
```



