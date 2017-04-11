var gulp = require ('gulp'),
    sass = require ('gulp-sass'),
    path1 = 'styles/**/*.scss';
gulp.task('default', function(){
    gulp.watch(path1, ['styles']);
});

gulp.task('styles', function(){
   gulp.src(path1)
       .pipe(sass().on('error', sass.logError))
       .pipe(gulp.dest('./styles'));
});