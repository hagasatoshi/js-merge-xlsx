/**
 * underscore mixin
 * utility functions for js-merge-xlsx.
 * @author Satoshi Haga
 * @date 2015/10/03
 */

const _ = require('underscore');
const Mustache = require('mustache');

_.mixin({
    isString: (arg)=>{
        return (typeof arg === 'string');
    },
    stringValue: (xml2jsElement)=>{
        if(!_.isArray(xml2jsElement)){
            return xml2jsElement;
        }
        if(xml2jsElement[0]._){
            return xml2jsElement[0]._;
        }
        return xml2jsElement[0];
    },
    variables: (template)=>{
        if(!_(template).isString()){
            return null;
        }
        return _.map( _.filter(Mustache.parse(template),(e)=>(e[0] === 'name')), (e)=> e[1]);
    },
    hasVariable: (template)=> {
        return _.isString(template) && (_.variables(template).length !== 0)
    },
    deepCopy: (obj)=>JSON.parse(JSON.stringify(obj)),
    deleteProperties: (data, properties)=>{
        let isArray = _.isArray(data);
        if(!isArray) data = [data];
        _.each(data, (e)=> _.each(properties, (prop)=>delete e[prop]));
        return isArray? data : data[0];
    },
    decode: (val)=>{
        if(!val || (typeof val !== 'string')) return val;
        let decodeMap = {'&lt;': '<', '&gt;': '>', '&quot;': '"', '&#39;': '\'', '&amp;': '&'};
        return val.replace(/(&lt;|&gt;|&quot;|&#39;|&amp;)/g, (str, item)=>decodeMap[item]);
    }
});
