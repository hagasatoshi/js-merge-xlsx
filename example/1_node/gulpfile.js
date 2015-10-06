/**
 * * Gulp task definition
 * * @author Satoshi Haga
 * * @date 2015/10/02
 **/
var gulp = require('gulp');
var babel = require('gulp-babel');

/* babel compile task */
gulp.task('babel', function () {

    //source JavaScript files
    gulp.src('src/app.js')
        .pipe(babel())
        .pipe(gulp.dest('./'));

});

/* default task */
gulp.task('default',['babel']);