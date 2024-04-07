'use strict';

import sh from 'shelljs';
import upath from 'upath';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

// Lấy đường dẫn file hiện tại và thư mục chứa file đó
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Hàm renderAssets để copy tài nguyên
export function renderAssets() {
    // copy assets
    let sourcePath = upath.resolve(__dirname, '../src/assets');
    let destPath = upath.resolve(__dirname, '../dist');

    // Sử dụng shelljs để copy thư mục
    sh.cp('-R', sourcePath, destPath);  
}
renderAssets();