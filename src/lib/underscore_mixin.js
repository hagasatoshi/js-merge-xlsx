/**
 * underscore mixin
 * utility functions for js-merge-xlsx.
 * @author Satoshi Haga
 * @date 2015/10/03
 */

const _ = require('underscore');
const Mustache = require('mustache');
const xml2js = require('xml2js');
const builder = new xml2js.Builder();

_.mixin({
    isString: (arg) => {
        return (typeof arg === 'string');
    },

    stringValue: (xml2jsElement) => {
        if(!_.isArray(xml2jsElement)) {
            return xml2jsElement;
        }
        if(xml2jsElement[0]._) {
            return xml2jsElement[0]._;
        }
        return xml2jsElement[0];
    },

    variables: (template) => {
        if(!_.isString(template)) {
            return null;
        }
        // TODO should return only element having variables as '{{}}' and '{{{}}}'
        // by regular expression
        //return _.map(_.filter(Mustache.parse(template), (e) => (e[0] === 'name')), (e) => e[1]);
        return _.map(template, (e) => e[1]);
    },

    hasVariable: (template) => {
        return _.isString(template) && (_.variables(template).length !== 0)
    },

    deepCopy: (obj) => JSON.parse(JSON.stringify(obj)),

    deleteProperties: (data, properties) => {
        let isArray = _.isArray(data);
        if(!isArray) data = [data];
        _.each(data, (e) => _.each(properties, (prop) => delete e[prop]));
        return isArray? data : data[0];
    },

    sum: (arrayObj, valueFn) => _.reduce(arrayObj, (sum, obj) => valueFn(obj), 0),

    count: (arrayObj, criteriaFn) => _.sum(arrayObj, (obj) => criteriaFn(obj) ? 1 : 0),

    reverseEach: (arrayObj, fn) => {
        _.each(_.sortBy(arrayObj, (obj, index) => (-1) * index), fn);
    },

    nestedEach: (array1, array2, fn) => {
        _.each(array1, (e1) => {
            _.each(array2, (e2) => {
                fn(e1, e2);
            });
        });
    },

    splice: (arrayObj, criteriaFn) => {
        _.reverseEach(arrayObj, (obj, index) => {
            if(criteriaFn(obj)) {
                arrayObj.splice(index, 1);
            }
        })
    },

    containsAsPartialString: (array, str) => {
        return _.reduce(array, (contained, e) => {
            return contained || (e.indexOf(str) !== -1);
        }, false)
    },

    consistOf: (elm, props) => {
        return _.reduce(props, (consist, prop) => {
            return consist && (elm[prop] !== undefined);
        }, true)
    },

    allConsistOf: (array, props) => {
        return _.reduce(array, (consist, elm) => {
            return consist && _.consistOf(elm, props);
        }, true);
    },

    includeString: (str, keyword) => {
        return str.indexOf(keyword) !== -1;
    }
});
