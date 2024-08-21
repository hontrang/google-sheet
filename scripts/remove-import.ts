import gulp from 'gulp';
import replace from 'gulp-replace';
import { Stream } from 'stream';

export function removeImports(): Stream {
  return gulp
    .src('dist/*.ts')
    .pipe(replace(/^(import\s.+?;)/gm, '//$1'))
    .pipe(gulp.dest('dist'));
}

removeImports();
