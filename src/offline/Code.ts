import { ExcelHelper } from "./ExcelHelper";


async function main() {
    const excel = new ExcelHelper();
    await excel.truyCapWorkBook();
    // const value = excel.ghiDuLieuVaoO("hose", "hose", "A1");
    const value = excel.xoaDuLieuTrongCot("DC", "C", 2, 5);
    // const value1 = excel.layDuLieuTrongCot("hose", "B");
    await excel.luuWorkBook();
}

main();

