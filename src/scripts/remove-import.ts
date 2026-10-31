import gulp from 'gulp';
import replace from 'gulp-replace';

export function removeImports(): NodeJS.ReadWriteStream {
  return gulp
    .src('dist/*.ts')
    .pipe(replace(/^(import\s.+?;)/gm, '//$1'))
    .pipe(gulp.dest('dist'));
}

removeImports();
