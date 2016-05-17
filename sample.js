var _ = require('underscore');

var add = function(e) {
    return e + 1;
};

var multiple = function(e) {
    return 2 * e;
};

var divide = function(e) {
    return e / 3;
};

var compose = _.compose(divide, multiple, add);
console.log(compose(1));