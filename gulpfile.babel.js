/**
 * Gulp task definition
 * @author Satoshi Haga
 * @date 2015/09/30
 */
const gulp = require('gulp');
const runSequence = require('run-sequence');
const babel = require('gulp-babel');
const uglify = require('gulp-uglify');
const mocha = require('gulp-mocha');

gulp.task('compile', ()=>{
    return gulp.src('src/**/*.js')
        .pipe(babel())
        .pipe(gulp.dest('./'));
});

gulp.task('compress', ['compress-excelmerge', 'compress-lib']);
gulp.task('compress-excelmerge', ()=>{
    return gulp.src('./excelmerge.js')
        .pipe(uglify())
        .pipe(gulp.dest('./'));
});
gulp.task('compress-lib', ()=>{
    return gulp.src('lib/*.js')
        .pipe(uglify())
        .pipe(gulp.dest('lib'));
});

gulp.task('test-setup', ()=>{
    return gulp.src('src_test/mocha/**/*.js')
        .pipe(babel())
        .pipe(gulp.dest('test/mocha'));
});

gulp.task('mocha', ()=>{
    return gulp.src('test/mocha/test.js')
        .pipe(mocha());
});

gulp.task('default',(cb)=>{
    runSequence(
        ['compile','test-setup'], 'compress', 'mocha', cb
    )
});

