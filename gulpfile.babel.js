/**
 * * Gulp task definition
 * * @author Satoshi Haga
 * * @date 2015/09/30
 **/
import gulp from 'gulp'
import babel from 'gulp-babel'
import mocha from 'gulp-mocha'
import runSequence from 'run-sequence'

/* babel compile task */
gulp.task('babel', ()=>{
    //source JavaScript files
    gulp.src('src/**/*.js')
        .pipe(babel())
        .pipe(gulp.dest('./'));
    //test JavaScript files
    gulp.src('src_test/**/*.js')
        .pipe(babel())
        .pipe(gulp.dest('test/'));
});

/* mocha testing task */
gulp.task('mocha', ()=>{
    gulp.src('test/test.js', {read: false})
        .pipe(mocha());
});


/* copy test resources */
gulp.task('copy_resources', ()=>{
    gulp.src('src_test/templates/*.xlsx')
        .pipe(gulp.dest('test/templates/'));
    gulp.src('src_test/input/*.*')
        .pipe(gulp.dest('test/input/'));

});

/* default task */
gulp.task('default', (callback)=> {
    runSequence(
        'babel',
        'copy_resources',
        'mocha',
        callback
    )
});