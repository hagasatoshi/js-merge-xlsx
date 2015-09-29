/**
 * * Gulp task definition
 * * @author Satoshi Haga
 * * @date 2015/09/30
 **/
var gulp = require('gulp');
var babel = require('gulp-babel');

/* babel compile task */
gulp.task('babel', function () {

    //source JavaScript files
    gulp.src('src/**/*.js')
        .pipe(babel())
        .pipe(gulp.dest('build'));

    //example JavaScript files
    gulp.src('example/**/*.js')
        .pipe(babel())
        .pipe(gulp.dest('build_example'));
});

/* default task */
gulp.task('default',['babel']);