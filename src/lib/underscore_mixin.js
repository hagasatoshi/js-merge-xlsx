/**
 * * underscore mixin
 * * utility functions for js-merge-xlsx.
 * * @author Satoshi Haga
 * * @date 2015/10/03
 **/

import _ from 'underscore'
import Mustache from 'mustache'

_.mixin({
    /**
     * * is_string
     * * @param {Object} arg
     * * @returns {boolean}
     */
    is_string: (arg)=>{
        return (typeof arg === 'string');
    },

    /**
     * * string_value
     * * return string value.
     * * @param arg
     * * @returns {String}
     */
    string_value: (xml2js_element)=>{
        if(!_.isArray(xml2js_element)){
            return xml2js_element;
        }
        if(xml2js_element[0]._){
            return xml2js_element[0]._;
        }
        return xml2js_element[0];
    },

    /**
     * * variables
     * * pick up and return the list of mustache-variables
     * * @param {String} template
     * * @returns {Array}
     */
    variables: (template)=>{
        if(!_(template).is_string()){
            return null;
        }
        return _.map( _.filter(Mustache.parse(template),(e)=>(e[0] === 'name')), (e)=> e[1]);
    },

    /**
     * * has_variable
     * * check if parameter-string has mustache-variables or not
     * * @param {String} template
     * * @returns {boolean}
     */
    has_variable: (template)=> {
        return _(template).is_string() && (_(template).variables().length !== 0)
    },

    //TODO this is temporary solution for lodash#deepCoy(). clarify why lodash#deepCoy() is so slow.
    /**
     * * deep_copy
     * * workaround for lodash#deepCoy(). clone object.
     * * @param {Object} obj
     * * @returns {Object}
     */
    deep_copy: (obj)=>JSON.parse(JSON.stringify(obj))
});