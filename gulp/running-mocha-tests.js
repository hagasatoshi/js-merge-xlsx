/**
 * Gulp task definition
 * Run mocha tests
 * @author Satoshi Haga
 * @date 2015/09/30
 */
var gulp = require('gulp');
var mocha = require('gulp-mocha');

gulp.task('running-mocha-tests', function(){
    return gulp.src('test/mocha/test.js')
        .pipe(mocha());
});