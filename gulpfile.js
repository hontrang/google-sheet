const gulp = require('gulp');
const exec = require('child_process').exec;

gulp.task('push',gulp.series(function() {
    exec('clasp push');
}));

gulp.task('watch', gulp.series(function() {
    gulp.watch([
    '**/*.gs'
], gulp.series('push'));
}));

gulp.task('default', gulp.series('watch'));