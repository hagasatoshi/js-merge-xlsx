/**
 * Gulp task definition
 * @author Satoshi Haga
 * @date 2015/09/30
 */
const gulp = require('gulp');
const $ = require('gulp-load-plugins')();
const runSequence = require('run-sequence');

const config = {
    js: {
        src: 'src/**/*.js',
        dest: './'
    },
    uglify: {
        src: './excelmerge.js',
        src_lib: 'lib/*.js',
        dest: './',
        dest_lib: 'lib'
    },
    test: {
        src: 'src_test/mocha/**/*.js',
        dest: 'test/mocha',
        main: 'test/mocha/test.js'
    }
};

gulp.task('wrk', ()=>{
    return gulp.src('./wrk/*.js')
        .pipe($.babel())
        .pipe(gulp.dest('./wrkBuild'));
});


gulp.task('compile', ()=>{
    return gulp.src(config.js.src)
        .pipe($.babel())
        .pipe(gulp.dest(config.js.dest));
});

gulp.task('compress', ['compress-excelmerge', 'compress-lib']);
gulp.task('compress-excelmerge', ()=>{
    return gulp.src(config.uglify.src)
        .pipe($.uglify())
        .pipe(gulp.dest(config.uglify.dest));
});
gulp.task('compress-lib', ()=>{
    return gulp.src(config.uglify.src_lib)
        .pipe($.uglify())
        .pipe(gulp.dest(config.uglify.dest_lib));
});

gulp.task('test-setup', ()=>{
    return gulp.src(config.test.src)
        .pipe($.babel())
        .pipe(gulp.dest(config.test.dest));
});

gulp.task('mocha', ()=>{
    return gulp.src(config.test.main)
        .pipe($.mocha());
});

gulp.task('default',(cb)=>{
    runSequence(
        ['compile','test-setup'], 'compress', 'mocha', cb
    )
});

