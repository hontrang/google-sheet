import sh from 'shelljs';
import upath from 'upath';

// Tạo đường dẫn đích
const destPath: string = upath.resolve(__dirname, '../../dist');
const buildPath: string = upath.resolve(__dirname, '../../build');

// Xóa tất cả các file trong thư mục đích
sh.rm('-rf', [`${destPath}/*`, `${buildPath}/*`]);
