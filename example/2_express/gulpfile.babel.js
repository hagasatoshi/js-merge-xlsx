const gulp = require('gulp');
const $ = require('gulp-load-plugins')();
const webpack = require('webpack-stream');
const runSequence = require('run-sequence');

const config = {
    js: {
        client: {
            src:  'src/js/client/**/*.js',
            dest: 'build'
        },
        server: {
            src:  'src/js/server/server.js',
            dest: './'
        }
    },
    sass: {
        src:  './src/styles/*.scss',
        dest: 'public/styles'
    },
    webpack: {
        src:  './build/app.js',
        file: 'js_merge_xlsx.js',
        dest: 'public/scripts/'
    }
};

gulp.task('babel-client', () => {
    return gulp.src(config.js.client.src)
        .pipe($.babel())
        .pipe(gulp.dest(config.js.client.dest));
});

gulp.task('babel-server', () => {
    return gulp.src(config.js.server.src)
        .pipe($.babel())
        .pipe(gulp.dest(config.js.server.dest));
});

gulp.task('sass', () => {
    return gulp.src(config.sass.src)
        .pipe($.sass())
        .pipe(gulp.dest(config.sass.dest));
});

gulp.task('webpack', () => {
    return gulp.src(config.webpack.src)
        .pipe(webpack({
            output: {
                filename: config.webpack.file
            }
        }))
        .pipe(gulp.dest(config.webpack.dest));
});

gulp.task('default', (cb) => {
    runSequence(
        ['babel-client','babel-server','sass'],
        'webpack',
        cb
    )
});