/**
 * * Gulp task definition
 * * @author Satoshi Haga
 * * @date 2015/09/30
 **/
import gulp from 'gulp'
import babel from 'gulp-babel'
import mocha from 'gulp-mocha'
import runSequence from 'run-sequence'

/* compile source resources */
gulp.task('babel-src', ()=>{
    return gulp.src('src/**/*.js')
        .pipe(babel())
        .pipe(gulp.dest('./'));
});

/* compile test resources */
gulp.task('babel-test', ()=>{
    return gulp.src('src_test/**/*.js')
        .pipe(babel())
        .pipe(gulp.dest('test/'));
});

/* mocha testing */
gulp.task('mocha', ()=>{
    return gulp.src('test/test.js')
        .pipe(mocha());
});

/* copy test resources(MS-Excel templates) */
gulp.task('copy_templates', ()=>{
    return gulp.src('src_test/templates/*.xlsx')
        .pipe(gulp.dest('test/templates/'));
});

/* copy test resources(yaml) */
gulp.task('copy_yaml_files', ()=>{
    return gulp.src('src_test/input/*.*')
        .pipe(gulp.dest('test/input/'));
});

/* default task */
gulp.task('default', (callback)=> {
    runSequence(
        ['babel-src','babel-test','copy_templates','copy_yaml_files'],
        'mocha',
        callback
    )
});