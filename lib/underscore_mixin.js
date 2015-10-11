/**
 * * underscore mixin
 * * utility functions for js-merge-xlsx.
 * * @author Satoshi Haga
 * * @date 2015/10/03
 **/

'use strict';

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { 'default': obj }; }

var _underscore = require('underscore');

var _underscore2 = _interopRequireDefault(_underscore);

var _mustache = require('mustache');

var _mustache2 = _interopRequireDefault(_mustache);

_underscore2['default'].mixin({
    /**
     * * is_string
     * * @param {Object} arg
     * * @returns {boolean}
     */
    is_string: function is_string(arg) {
        return typeof arg === 'string';
    },

    /**
     * * string_value
     * * return string value.
     * * @param arg
     * * @returns {String}
     */
    string_value: function string_value(xml2js_element) {
        if (!_underscore2['default'].isArray(xml2js_element)) {
            return xml2js_element;
        }
        if (xml2js_element[0]._) {
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
    variables: function variables(template) {
        if (!(0, _underscore2['default'])(template).is_string()) {
            return null;
        }
        return _underscore2['default'].map(_underscore2['default'].filter(_mustache2['default'].parse(template), function (e) {
            return e[0] === 'name';
        }), function (e) {
            return e[1];
        });
    },

    /**
     * * has_variable
     * * check if parameter-string has mustache-variables or not
     * * @param {String} template
     * * @returns {boolean}
     */
    has_variable: function has_variable(template) {
        return (0, _underscore2['default'])(template).is_string() && (0, _underscore2['default'])(template).variables().length !== 0;
    },

    //TODO this is temporary solution for lodash#deepCoy(). clarify why lodash#deepCoy() is so slow.
    /**
     * * deep_copy
     * * workaround for lodash#deepCoy(). clone object.
     * * @param {Object} obj
     * * @returns {Object}
     */
    deep_copy: function deep_copy(obj) {
        return JSON.parse(JSON.stringify(obj));
    }
});