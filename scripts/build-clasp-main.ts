import sh from 'shelljs';
import upath from 'upath';
import { fileURLToPath } from 'url';
import { dirname } from 'path';
import glob from 'glob';

const __filename: string = fileURLToPath(import.meta.url);
const __dirname: string = dirname(__filename);

export function renderClasp(): void {
  // copy clasp
  const sourcePath: string = upath.resolve(__dirname, '/Users/hontrang/code/clasp/main');
  const destPath: string = upath.resolve(__dirname, '../dist');
  console.log('Sao chép files trong folder clasp vào dist');

  // Lấy tất cả file trong thư mục nguồn, bao gồm cả file ẩn
  glob('**/*', { dot: true, cwd: sourcePath }, function (er: Error | null, files: string[]): void {
    if (er) {
      console.error('Lỗi khi tìm file: ', er);
      return;
    }

    // Sao chép từng file
    files.forEach((file: string) => {
      const srcFilePath: string = upath.join(sourcePath, file);
      const destFilePath: string = upath.join(destPath, file);
      sh.cp('-Rf', srcFilePath, destFilePath);
    });

    console.log('Sao chép thành công.');
  });
}

renderClasp();