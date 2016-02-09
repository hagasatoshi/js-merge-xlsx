/**
 * Gulp task definition
 * Compile Source Resources
 * @author Satoshi Haga
 * @date 2015/09/30
 */
var gulp = require('gulp');
var babel = require('gulp-babel');


gulp.task('compile-source-resources', function(){
    return gulp.src('src/**/*.js')
        .pipe(babel())
        .pipe(gulp.dest('./'));
});
