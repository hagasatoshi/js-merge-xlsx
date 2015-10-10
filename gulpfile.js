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
        .pipe(gulp.dest('./'));
    //test JavaScript files
    gulp.src('src_test/**/*.js')
        .pipe(babel())
        .pipe(gulp.dest('test/'));
});

/* copy test templates */
gulp.task('copy_templates', function () {
    gulp.src('src_test/templates/*.xlsx')
        .pipe(gulp.dest('test/templates/'));
});

/* default task */
gulp.task('default',['babel','copy_templates']);