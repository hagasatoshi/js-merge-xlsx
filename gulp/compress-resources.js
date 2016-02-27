/**
 * Gulp task definition
 * Compile Source Resources
 * @author Satoshi Haga
 * @date 2015/09/30
 */
var gulp = require('gulp');
var uglify = require('gulp-uglify');

gulp.task('compress', ['compress-resources', 'compress-resources-lib']);

gulp.task('compress-resources', function(){
    return gulp.src('./excelmerge.js')
        .pipe(uglify())
        .pipe(gulp.dest('./'));
});

gulp.task('compress-resources-lib', function(){
    return gulp.src('lib/*.js')
        .pipe(uglify())
        .pipe(gulp.dest('lib'));
});
