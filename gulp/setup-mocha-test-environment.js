/**
 * * Gulp task definition
 * * Setup mocha test environment
 * * @author Satoshi Haga
 * * @date 2015/09/30
 **/
var gulp = require('gulp');
var babel = require('gulp-babel');

gulp.task('setup-mocha-test-environment', function(){
    return gulp.src('src_test/mocha/**/*.js')
        .pipe(babel())
        .pipe(gulp.dest('test/mocha'));
});

