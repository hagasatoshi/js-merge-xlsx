const gulp = require('gulp');
const babel = require('gulp-babel');

gulp.task('babel', ()=>{
    return gulp.src('src/app.js')
        .pipe(babel())
        .pipe(gulp.dest('./'));
});

gulp.task('default',['babel']);