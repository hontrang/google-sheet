import sh from 'shelljs';
import upath from 'upath';
import { fileURLToPath } from 'url';
import { dirname } from 'path';
import glob from 'glob';

const fileName: string = fileURLToPath(import.meta.url);
const dirName: string = dirname(fileName);

export function renderClasp(): void {
  // copy clasp
  const sourcePath: string = upath.resolve(dirName, process.env.dwClasp);
  const destPath: string = upath.resolve(dirName, '../../dist');
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
