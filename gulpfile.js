const gulp = require('gulp');
const exec = require('child_process').exec;

gulp.task('push', function (done) {
    exec('clasp push', function (err,stdout) {
        console.log(stdout);
    });
    done();
});

gulp.task('watch', function () {
    gulp.watch('**/*.js', gulp.series('push'));
});

gulp.task('default', gulp.series('watch'));