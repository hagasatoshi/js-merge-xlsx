/**
 * * Gulp task definition
 * * @author Satoshi Haga
 * * @date 2015/09/30
 **/
var gulp = require('gulp');
var babel = require('gulp-babel');
var sass = require('gulp-sass');
var webpack = require('webpack-stream');
var runSequence = require('run-sequence');


/* babel compile task */
gulp.task('babel', function () {

    //client resources
    gulp.src('src/js/client/**/*.js')
        .pipe(babel())
        .pipe(gulp.dest('build'));

    //server resources
    gulp.src('src/js/server/server.js')
        .pipe(babel())
        .pipe(gulp.dest('./'));

});

/* sass compile task */
gulp.task('sass', function(){
    gulp.src('./src/styles/*.scss')
        .pipe(sass())
        .pipe(gulp.dest('public/styles'));
});

/* task building 'js_merge_xlsx.js' */
gulp.task('webpack', function() {
    gulp.src('./build/app.js')
        .pipe(webpack({
            output: {
                filename: 'js_merge_xlsx.js'
            }
        }))
        .pipe(gulp.dest('public/scripts/'));
});

/* default task */
gulp.task('default', function(callback) {
    runSequence(
        'babel',
        'sass',
        'webpack',
        callback
    )
});