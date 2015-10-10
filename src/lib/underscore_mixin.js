/**
 * * underscore mixin
 * * @author Satoshi Haga
 * * @date 2015/10/03
 **/

import _ from 'underscore'
import Mustache from 'mustache'

_.mixin({
    is_string: (arg)=>{
        return (typeof arg === 'string');
    },
    string_value: (xml2js_element)=>{
        if(!_.isArray(xml2js_element)){
            return xml2js_element;
        }
        if(xml2js_element[0]._){
            return xml2js_element[0]._;
        }
        return xml2js_element[0];
    },
    variables: (template)=>{
        if(!_(template).is_string()){
            return null;
        }
        return _.map( _.filter(Mustache.parse(template),(e)=>(e[0] === 'name')), (e)=> e[1]);
    },
    has_variable: (template)=> {
        return _(template).is_string() && (_(template).variables().length !== 0)
    },
    //TODO this is temporary solution for lodash#deepCoy(). clarify why lodash#deepCoy() is so slow.
    deep_copy: (obj)=>JSON.parse(JSON.stringify(obj))
});