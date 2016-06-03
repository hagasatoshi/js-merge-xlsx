/**
 * underscore mixin
 * utility functions for js-merge-xlsx.
 * @author Satoshi Haga
 * @date 2015/10/03
 */

'use strict';

var _ = require('underscore');
var Mustache = require('mustache');
var xml2js = require('xml2js');
var builder = new xml2js.Builder();

_.mixin({
    stringValue: function stringValue(elm) {
        if (!_.isArray(elm)) return elm;

        elm = elm[0];
        return elm._ ? elm._ : elm;
    },

    variables: function variables(template) {
        if (!_.isString(template)) {
            return null;
        }
        return _.map(_.filter(Mustache.parse(template), function (e) {
            return e[0] === '&';
        }), function (e) {
            return e[1];
        });
    },

    hasVariable: function hasVariable(template) {
        return _.isString(template) && _.variables(template).length !== 0;
    },

    deepCopy: function deepCopy(obj) {
        if (!_.isObject(obj)) {
            throw new Error("_.deepCopy() : argument should be object.");
        }
        return JSON.parse(JSON.stringify(obj));
    },

    deleteProps: function deleteProps(data, properties) {
        var recursive = function recursive(arrayObj, props) {
            return _.reduce(arrayObj, function (array, elm) {
                array.push(_.deleteProps(elm, props));
                return array;
            }, []);
        };
        return _.isArray(data) ? recursive(data, properties) : _.reduce(properties, function (obj, prop) {
            delete obj[prop];
            return obj;
        }, data);
    },

    sum: function sum(arrayObj, valueFn) {
        return _.reduce(arrayObj, function (sum, obj) {
            return sum + valueFn(obj);
        }, 0);
    },

    count: function count(arrayObj, criteriaFn) {
        return _.sum(arrayObj, function (obj) {
            return criteriaFn(obj) ? 1 : 0;
        });
    },

    reverseEach: function reverseEach(arrayObj, fn) {
        _.each(_.chain(arrayObj).reverse().value(), fn);
    },

    arrayFrom: function arrayFrom(length) {
        return Array.apply(null, { length: length }).map(Number.call, Number);
    },

    //non-destructive for arrayObj
    reduceInReverse: function reduceInReverse(arrayObj, fn, initialValue) {
        var indexes = _.arrayFrom(arrayObj.length);
        indexes = _.chain(indexes).reverse().value();

        return _.reduce(indexes, function (x, index) {
            return fn(x, arrayObj[index], index);
        }, initialValue);
    },

    nestedEach: function nestedEach(array1, array2, fn) {
        _.each(array1, function (e1) {
            _.each(array2, function (e2) {
                fn(e1, e2);
            });
        });
    },

    //destructive change for arrayObj
    splice: function splice(arrayObj, criteriaFn) {
        return _.reduceInReverse(arrayObj, function (array, elm, index) {
            if (criteriaFn(elm)) {
                array.splice(index, 1);
            }
            return array;
        }, arrayObj);
    },

    containsAsPartial: function containsAsPartial(array, str) {
        return _.reduce(array, function (contained, e) {
            return contained || _.includeString(e, str);
        }, false);
    },

    consistOf: function consistOf(obj, props) {
        if (_.isArray(obj)) {
            return _.reduce(obj, function (consist, e) {
                return consist && _.consistOf(e, props);
            }, true);
        }
        if (_.isString(props)) {
            return _.has(obj, props);
        }
        if (_.isArray(props)) {
            return _.reduce(props, function (consist, prop) {
                return consist && _.consistOf(obj, prop);
            }, true);
        }
        return _.reduce(props, function (consist, prop, key) {
            return consist && obj[key] && _.consistOf(obj[key], prop);
        }, true);
    },

    includeString: function includeString(str, keyword) {
        return !!keyword && str.indexOf(keyword) !== -1;
    }
});