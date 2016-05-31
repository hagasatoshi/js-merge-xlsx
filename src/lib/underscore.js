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
    stringValue: (elm) => {
        if(!_.isArray(elm))
            return elm;

        elm = elm[0];
        return elm._ ? elm._ : elm;
    },

    variables: (template) => {
        if(!_.isString(template)) {
            return null;
        }
        return _.map(_.filter(Mustache.parse(template), (e) => (e[0] === '&')), (e) => e[1]);
    },

    hasVariable: (template) => {
        return _.isString(template) && (_.variables(template).length !== 0)
    },

    deepCopy: (obj) => {
        if(!_.isObject(obj)) {
            throw new Error("_.deepCopy() : argument should be object.");
        }
        return JSON.parse(JSON.stringify(obj));
    },

    deleteProps: (data, properties) => {
        const recursive = (arrayObj, props) => {
            return _.reduce(arrayObj, (array, elm) => {
                array.push(_.deleteProps(elm, props));
                return array;
            }, []);
        };
        return _.isArray(data) ? recursive(data, properties) :
            _.reduce(properties, (obj, prop) => {
                delete obj[prop];
                return obj;
            }, data);
    },

    sum: (arrayObj, valueFn) => _.reduce(arrayObj, (sum, obj) => sum + valueFn(obj), 0),

    count: (arrayObj, criteriaFn) => _.sum(arrayObj, (obj) => criteriaFn(obj) ? 1 : 0),

    reverseEach: (arrayObj, fn) => {
        _.each(
            _.chain(arrayObj).reverse().value(), fn
        );
    },

    arrayFrom: (length) => {
        return Array.apply(null, {length}).map(Number.call, Number);
    },

    //non-destructive for arrayObj
    reduceInReverse: (arrayObj, fn, initialValue) => {
        let indexes = _.arrayFrom(arrayObj.length);
        indexes = _.chain(indexes).reverse().value();

        return _.reduce(indexes, (x, index) => {
            return fn(x, arrayObj[index]);
        }, initialValue);
    },

    nestedEach: (array1, array2, fn) => {
        _.each(array1, (e1) => {
            _.each(array2, (e2) => {
                fn(e1, e2);
            });
        });
    },

    splice: (arrayObj, criteriaFn) => {
        return _.reduceInReverse(arrayObj, (array, elm, index) => {
            if(criteriaFn(elm)) {
                array.splice(index, 1);
            }
            return array;
        }, arrayObj);
    },

    containsAsPartialString: (array, str) => {
        return _.reduce(array, (contained, e) => {
            return contained || (e.indexOf(str) !== -1);
        }, false)
    },

    consistOf: (obj, props) => {
        if(_.isArray(obj)) {
            return _.reduce(obj, (consist, e) => {
                return consist && _.consistOf(e, props);
            }, true);
        }
        if(_.isString(props)) {
            return _.has(obj, props);
        }
        if(_.isArray(props)) {
            return _.reduce(props, (consist, prop) => {
                return consist && _.consistOf(obj, prop);
            }, true);
        }
        return _.reduce(props, (consist, prop, key) => {
            return consist && obj[key] &&
                _.consistOf(obj[key], prop);
        }, true);
    },

    includeString: (str, keyword) => {
        return str.indexOf(keyword) !== -1;
    }
});
