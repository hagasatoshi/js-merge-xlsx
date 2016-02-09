/**
 * underscore mixin
 * utility functions for js-merge-xlsx.
 * @author Satoshi Haga
 * @date 2015/10/03
 */

'use strict';

var _ = require('underscore');
var Mustache = require('mustache');
var builder = require('xml2js').Builder();

_.mixin({
    /**
     * isString
     * @param {Object} arg
     * @returns {boolean}
     */
    isString: function isString(arg) {
        return typeof arg === 'string';
    },

    /**
     * stringValue
     * return string value.
     * @param arg
     * @returns {String}
     */
    stringValue: function stringValue(xml2jsElement) {
        if (!_.isArray(xml2jsElement)) {
            return xml2jsElement;
        }
        if (xml2jsElement[0]._) {
            return xml2jsElement[0]._;
        }
        return xml2jsElement[0];
    },

    /**
     * variables
     * pick up and return the list of mustache-variables
     * @param {String} template
     * @returns {Array}
     */
    variables: function variables(template) {
        if (!_(template).isString()) {
            return null;
        }
        return _.map(_.filter(Mustache.parse(template), function (e) {
            return e[0] === 'name';
        }), function (e) {
            return e[1];
        });
    },

    /**
     * hasVariable
     * check if parameter-string has mustache-variables or not
     * @param {String} template
     * @returns {boolean}
     */
    hasVariable: function hasVariable(template) {
        return _(template).isString() && _(template).variables().length !== 0;
    },

    //TODO this is temporary solution for lodash#deepCoy(). clarify why lodash#deepCoy() is so slow.
    /**
     * deepCopy
     * workaround for lodash#deepCoy(). clone object.
     * @param {Object} obj
     * @returns {Object}
     */
    deepCopy: function deepCopy(obj) {
        return JSON.parse(JSON.stringify(obj));
    },

    /**
     * deleteProperties
     * delete properties
     * @param {Object} data
     * @returns {Object}
     */
    deleteProperties: function deleteProperties(data, properties) {
        var isArray = _.isArray(data);
        if (!isArray) data = [data];
        _.each(data, function (e) {
            return _.each(properties, function (prop) {
                return delete e[prop];
            });
        });
        return isArray ? data : data[0];
    },

    /**
     * decode
     * @param {String} val
     * @returns {String}
     */
    decode: function decode(val) {
        if (!val || typeof val !== 'string') return val;
        var decodeMap = { '&lt;': '<', '&gt;': '>', '&quot;': '"', '&#39;': '\'', '&amp;': '&' };
        return val.replace(/(&lt;|&gt;|&quot;|&#39;|&amp;)/g, function (str, item) {
            return decodeMap[item];
        });
    }

});