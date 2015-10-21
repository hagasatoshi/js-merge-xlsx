/**
 * * underscore mixin
 * * utility functions for js-merge-xlsx.
 * * @author Satoshi Haga
 * * @date 2015/10/03
 **/

var _ = require('underscore');
var Mustache = require('mustache');
var builder = require('xml2js').Builder();

_.mixin({
    /**
     * * isString
     * * @param {Object} arg
     * * @returns {boolean}
     */
    isString: (arg)=>{
        return (typeof arg === 'string');
    },

    /**
     * * stringValue
     * * return string value.
     * * @param arg
     * * @returns {String}
     */
    stringValue: (xml2jsElement)=>{
        if(!_.isArray(xml2jsElement)){
            return xml2jsElement;
        }
        if(xml2jsElement[0]._){
            return xml2jsElement[0]._;
        }
        return xml2jsElement[0];
    },

    /**
     * * variables
     * * pick up and return the list of mustache-variables
     * * @param {String} template
     * * @returns {Array}
     */
    variables: (template)=>{
        if(!_(template).isString()){
            return null;
        }
        return _.map( _.filter(Mustache.parse(template),(e)=>(e[0] === 'name')), (e)=> e[1]);
    },

    /**
     * * hasVariable
     * * check if parameter-string has mustache-variables or not
     * * @param {String} template
     * * @returns {boolean}
     */
    hasVariable: (template)=> {
        return _(template).isString() && (_(template).variables().length !== 0)
    },

    //TODO this is temporary solution for lodash#deepCoy(). clarify why lodash#deepCoy() is so slow.
    /**
     * * deepCopy
     * * workaround for lodash#deepCoy(). clone object.
     * * @param {Object} obj
     * * @returns {Object}
     */
    deepCopy: (obj)=>JSON.parse(JSON.stringify(obj)),

    /**
     * * decode
     * * @param {String} val
     * * @returns {String}
     */
    decode: (val)=>{
        if(!val || (typeof val !== 'string')) return val;
        let decodeMap = {'&lt;': '<', '&gt;': '>', '&quot;': '"', '&#39;': '\'', '&amp;': '&'};
        return val.replace(/(&lt;|&gt;|&quot;|&#39;|&amp;)/g, (str, item)=>decodeMap[item]);
    }

});
