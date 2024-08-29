import sh from 'shelljs';
import upath from 'upath';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

// Lấy đường dẫn của file hiện tại và thư mục chứa file đó
const __filename: string = fileURLToPath(import.meta.url);
const __dirname: string = dirname(__filename);

// Tạo đường dẫn đích
const destPath: string = upath.resolve(__dirname, '../dist');
const buildPath: string = upath.resolve(__dirname, '../build');

// Xóa tất cả các file trong thư mục đích
sh.rm('-rf', [`${destPath}/*`, `${buildPath}/*`]);
