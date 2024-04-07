'use strict';

import sh from 'shelljs';
import upath from 'upath';
import { fileURLToPath } from 'url';
import { dirname } from 'path';
import glob from 'glob';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

export function renderClasp() {
    // copy clasp
    let sourcePath = upath.resolve(__dirname, '../src/clasp');
    let destPath = upath.resolve(__dirname, '../dist');
    console.log("Sao chép files trong folder clasp vào dist");
    // Lấy tất cả file trong thư mục nguồn, bao gồm cả file ẩn
    glob("**/*", { dot: true, cwd: sourcePath }, function (er, files) {
        if (er) {
            console.error("Lỗi khi tìm file: ", er);
            return;
        }

        // Sao chép từng file
        files.forEach(file => {
            let srcFilePath = upath.join(sourcePath, file);
            let destFilePath = upath.join(destPath, file);
            sh.cp('-Rf', srcFilePath, destFilePath);
        });

        console.log("Sao chép thành công.");
    });
}
renderClasp();