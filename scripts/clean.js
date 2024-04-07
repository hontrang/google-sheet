'use strict';

// Sử dụng import thay vì require
import sh from 'shelljs';
import upath from 'upath';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

// Lấy đường dẫn của file hiện tại và thư mục chứa file đó
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Tạo đường dẫn đích
const destPath = upath.resolve(__dirname, '../dist');

// Xóa tất cả các file trong thư mục đích
sh.rm('-rf', `${destPath}/*`);