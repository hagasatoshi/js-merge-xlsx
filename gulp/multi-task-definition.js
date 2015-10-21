/**
 * * Gulp task definition
 * * multi task definition
 * * @author Satoshi Haga
 * * @date 2015/09/30
 **/
var gulp = require('gulp');
var requireDir = require('require-dir');
var runSequence = require('run-sequence');

/* default task */
gulp.task('default', function(callback){
    runSequence(
        ['compile-source-resources','setup-mocha-test-environment'],
        'running-mocha-tests',
        callback
    )
});

