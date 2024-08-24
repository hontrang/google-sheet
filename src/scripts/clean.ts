import sh from 'shelljs';
import upath from 'upath';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

// Lấy đường dẫn của file hiện tại và thư mục chứa file đó
const fileName: string = fileURLToPath(import.meta.url);
const dirName: string = dirname(fileName);

// Tạo đường dẫn đích
const destPath: string = upath.resolve(dirName, '../../dist');
const buildPath: string = upath.resolve(dirName, '../../build');

// Xóa tất cả các file trong thư mục đích
sh.rm('-rf', [`${destPath}/*`, `${buildPath}/*`]);
