// Import các module sử dụng ES Modules syntax
import gulp from 'gulp';
import replace from 'gulp-replace';

// Task để xóa các dòng chứa từ khóa `import`
export function removeImports() {
  return gulp.src('dist/*.js') // Thay đổi 'src/**/*.js' tùy vào cấu trúc thư mục của bạn
    .pipe(replace(/import\s.+?;/g, '')) // Sử dụng biểu thức chính quy để tìm và xóa
    .pipe(gulp.dest('dist')); // Lưu các file đã được xử lý vào thư mục 'dist'
}

removeImports();